import { useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import {
  createWorkbookReviewController,
  workbookReviewScopeKey,
  type WorkbookReviewFrame,
  type WorkbookSpellingFinding,
  type WorkbookReviewMode,
} from './workbookReviewController';

export interface WorkbookReviewOpenIntent {
  serial: number;
  mode: WorkbookReviewMode;
}
export interface WorkbookReviewProps {
  frame: WorkbookReviewFrame;
  visible?: boolean;
  openIntent?: WorkbookReviewOpenIntent;
  words?: (text: string) => string;
}
const identity = (text: string) => text;
const total = (values: Record<string, number>): number => Object.values(values).reduce((sum, value) => sum + value, 0);
export function WorkbookReview({
  frame,
  visible = true,
  openIntent,
  words = identity,
}: WorkbookReviewProps): JSX.Element {
  const controller = useMemo(() => createWorkbookReviewController(), []),
    snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot),
    key = workbookReviewScopeKey(frame.scope);
  const lifetime = useRef(0),
    lastOpenSerial = useRef(-1),
    [changes, setChanges] = useState<Record<string, string>>({});
  const fresh = !!(visible && frame.scope.active && snapshot.scope && workbookReviewScopeKey(snapshot.scope) === key),
    stats = fresh ? snapshot.statistics : null,
    findings = fresh ? snapshot.findings : [],
    busy = snapshot.physicalBusy,
    accepting = snapshot.phase === 'accepting';
  useLayoutEffect(() => {
    controller.setFrame({ ...frame, scope: { ...frame.scope, active: frame.scope.active && visible } });
    setChanges({});
    return () => {
      void controller.cancel();
    };
  }, [key, visible]);
  const current = (): boolean => {
    const scope = controller.getSnapshot().scope;
    return !!(visible && frame.scope.active && scope && workbookReviewScopeKey(scope) === key);
  };
  useLayoutEffect(() => {
    if (!openIntent || !current()) return;
    if (
      Object.keys(openIntent).sort().join(',') !== 'mode,serial' ||
      !Number.isSafeInteger(openIntent.serial) ||
      openIntent.serial < 0 ||
      !['statistics', 'spelling'].includes(openIntent.mode)
    )
      throw Error('Invalid Workbook Review open intent.');
    if (openIntent.serial <= lastOpenSerial.current) return;
    lastOpenSerial.current = openIntent.serial;
    controller.setMode(openIntent.mode);
  }, [openIntent?.serial, openIntent?.mode, visible]);
  useLayoutEffect(() => {
    const serial = ++lifetime.current;
    return () => {
      queueMicrotask(() => {
        if (lifetime.current === serial) void controller.dispose();
      });
    };
  }, []);
  const run = (job: Promise<unknown>): void => {
    void job.catch(() => {});
  };
  const change = (entry: WorkbookSpellingFinding): void => {
    if (current()) run(controller.changeWord(entry, changes[entry.id] ?? entry.suggestions[0] ?? ''));
  };
  const completed = fresh && (snapshot.mode === 'statistics' ? snapshot.statisticsComplete : snapshot.spellingComplete);
  const summary =
    !frame.scope.active || !visible
      ? 'Review is paused. Activate this workbook in Normal view to continue.'
      : busy
        ? accepting
          ? 'The original workbook is accepting the correction…'
          : snapshot.phase === 'idle'
            ? 'The previous inspection is stopping. Wait to scan this workbook.'
            : 'Reading the captured workbook…'
        : completed
          ? snapshot.mode === 'statistics'
            ? 'Review complete for the captured accepted records.'
            : 'Spelling scan finished for eligible words. Excluded records and tokens were not checked.'
          : 'Review incomplete.';
  return (
    <section
      className="sheet-review"
      aria-label={words('Workbook Review')}
      data-formula-draft-zone
      data-review-document={frame.scope.documentId}
      data-review-owner={frame.scope.ownerId}
      data-review-state={frame.scope.state}
      data-review-mode={snapshot.mode}
    >
      <div role="group" aria-label={words('Review mode')}>
        <button
          type="button"
          aria-pressed={snapshot.mode === 'statistics'}
          disabled={accepting}
          onClick={() => {
            if (current()) controller.setMode('statistics');
          }}
        >
          {words('Workbook Statistics')}
        </button>
        <button
          type="button"
          aria-pressed={snapshot.mode === 'spelling'}
          disabled={accepting}
          onClick={() => {
            if (current()) controller.setMode('spelling');
          }}
        >
          {words('Spelling')}
        </button>
      </div>
      <p>
        {words(
          'Inspects accepted records across all worksheets, including hidden rows and sheets. Blank worksheet positions without stored records are not counted.'
        )}
      </p>
      {snapshot.mode === 'spelling' && (
        <>
          <label>
            {words('Spelling language')}{' '}
            <select
              aria-label={words('Review locale')}
              value={snapshot.locale}
              disabled={accepting}
              onChange={(event) => {
                if (current()) {
                  setChanges({});
                  controller.setLocale(event.target.value as 'en-US' | 'en-GB');
                }
              }}
            >
              <option value="en-US">English (United States)</option>
              <option value="en-GB">English (United Kingdom)</option>
            </select>
          </label>
          <p>
            {words(
              'Checks literal text only. Formulas, numeric values, cached results and protected or special output cells are excluded. Suggestions do not change the workbook.'
            )}
          </p>
        </>
      )}
      <div>
        <button
          type="button"
          disabled={!frame.scope.active || !visible || busy || completed}
          onClick={() => {
            if (current())
              run(snapshot.mode === 'statistics' ? controller.scanStatistics() : controller.scanSpelling());
          }}
        >
          {words(
            snapshot.mode === 'statistics'
              ? snapshot.statisticsPages
                ? 'Continue statistics scan'
                : 'Scan workbook statistics'
              : snapshot.wordsChecked || snapshot.spellingRecords
                ? 'Continue spelling review'
                : 'Scan spelling'
          )}
        </button>
        <button
          type="button"
          disabled={!frame.scope.active || !visible || accepting}
          onClick={() => {
            if (current()) {
              setChanges({});
              controller.restart();
            }
          }}
        >
          {words('Restart review')}
        </button>
        {busy && !accepting && (
          <button
            type="button"
            onClick={() => {
              if (current()) run(controller.cancel());
            }}
          >
            {words('Cancel review')}
          </button>
        )}
      </div>
      <p role="status" aria-label={words('Review status')}>
        {words(summary)}
      </p>
      {fresh && snapshot.incompleteReason && <p role="status">{words(snapshot.incompleteReason)}</p>}
      {fresh && snapshot.error && <p role="alert">{words(snapshot.error)}</p>}
      {fresh && snapshot.notice && <p role="status">{words(snapshot.notice)}</p>}
      {snapshot.mode === 'statistics' && stats && (
        <>
          <p>
            {words(
              'Counts of records without formulas can include cached output records. Current, stored and pending results are counted separately. A read-only workbook does not mean each record has an explicit protection marker.'
            )}
          </p>
          <p>
            {snapshot.statisticsSheets} / {frame.scope.sheets.length} {words('worksheets completely scanned')} ·{' '}
            {snapshot.statisticsPages} {words('accepted-cell pages')}
          </p>
          <table aria-label={words('Workbook statistics')}>
            <tbody>
              {[
                ['Accepted records', stats.authoredCells],
                ['Authored blank records', stats.authoredBlankCells],
                ['Text records without formulas', stats.literalKinds.text],
                ['Number records without formulas', stats.literalKinds.number],
                ['Boolean records without formulas', stats.literalKinds.boolean],
                ['Error records without formulas', stats.literalKinds.error],
                ['Formula cells', stats.formulaCells],
                ['Current formula results', stats.formulaStates.calculated],
                ['Current formula errors', stats.formulaStates.error],
                ['Stored formula results', stats.formulaStates.stored],
                ['Pending formulas', stats.formulaStates.pending],
                ['Unresolved formulas', stats.formulaStates.unresolved],
                ['Circular formulas', stats.formulaStates.circular],
                ['Unqualified formulas', stats.formulaStates.unqualified],
                [
                  'Stored or pending results, including non-formula outputs',
                  total(stats.results.stored) + total(stats.results.pending),
                ],
                ['Special output records', stats.specialCells],
                ['Records with explicit protection markers', stats.protectedCells],
                ['Protected numeric precision', stats.unsafeNumericCells],
              ].map(([label, value]) => (
                <tr key={label}>
                  <th scope="row">{words(String(label))}</th>
                  <td>
                    <output aria-label={words(String(label))}>{value}</output>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}
      {snapshot.mode === 'spelling' && fresh && (
        <>
          <p>
            {snapshot.spellingSheets} / {frame.scope.sheets.length} {words('worksheets completely scanned')} ·{' '}
            {snapshot.spellingRecords} {words('accepted records examined')} · {snapshot.spellingTextCells}{' '}
            {words('literal text cells')} · {snapshot.wordsChecked} {words('words checked')} · {snapshot.findingsFound}{' '}
            {words('findings so far')}
          </p>
          <p>
            {snapshot.excludedCells} {words('ineligible records')} · {snapshot.excludedTokens}{' '}
            {words('oversized or numeric/identifier tokens excluded')}
          </p>
          {snapshot.availability && (
            <p aria-label={words('Offline dictionary availability')}>
              {snapshot.availability.locales
                .map(
                  (item) =>
                    `${item.locale}: ${item.status === 'available' && item.verified ? 'loaded ' + item.verified.package + ' ' + item.verified.version : 'unavailable'}`
                )
                .join(' · ')}
            </p>
          )}
          {!frame.scope.writable && (
            <p role="status">
              {words('This workbook is read-only. Findings can be inspected; corrections are disabled.')}
            </p>
          )}
          <ol aria-label={words('Spelling findings')}>
            {findings.map((entry) => (
              <li
                key={entry.id}
                data-review-cell={entry.finding.target.sheet + '!' + entry.finding.target.cell}
                data-review-token-start={entry.finding.token.start}
              >
                <strong>
                  {entry.finding.target.sheet}!{entry.finding.target.cell}
                </strong>{' '}
                <output aria-label={words('Original spelling word')}>{entry.finding.token.text}</output>
                <p>
                  {words('Original text positions')} {entry.finding.token.start}–{entry.finding.token.end}
                </p>
                <label>
                  {words('Replace with')}{' '}
                  <input
                    aria-label={
                      words('Replacement for') +
                      ' ' +
                      entry.finding.target.sheet +
                      '!' +
                      entry.finding.target.cell +
                      ' ' +
                      entry.finding.token.start
                    }
                    maxLength={512}
                    value={changes[entry.id] ?? entry.suggestions[0] ?? ''}
                    disabled={busy || !frame.scope.writable}
                    onChange={(event) => {
                      if (current()) setChanges((value) => ({ ...value, [entry.id]: event.target.value }));
                    }}
                  />
                </label>
                {entry.suggestions.length > 0 && (
                  <select
                    aria-label={
                      words('Suggestions for') +
                      ' ' +
                      entry.finding.target.sheet +
                      '!' +
                      entry.finding.target.cell +
                      ' ' +
                      entry.finding.token.start
                    }
                    disabled={busy || !frame.scope.writable}
                    value={changes[entry.id] ?? entry.suggestions[0]}
                    onChange={(event) => {
                      if (current()) setChanges((value) => ({ ...value, [entry.id]: event.target.value }));
                    }}
                  >
                    {entry.suggestions.map((value) => (
                      <option key={value} value={value}>
                        {value}
                      </option>
                    ))}
                  </select>
                )}
                <button
                  type="button"
                  disabled={busy || !frame.scope.writable || !(changes[entry.id] ?? entry.suggestions[0])}
                  title={!frame.scope.writable ? words('This workbook is read-only.') : undefined}
                  aria-label={
                    words('Change word in') +
                    ' ' +
                    entry.finding.target.sheet +
                    '!' +
                    entry.finding.target.cell +
                    ' ' +
                    entry.finding.token.start
                  }
                  onClick={() => change(entry)}
                >
                  {words('Change word')}
                </button>
                <button
                  type="button"
                  disabled={busy}
                  onClick={() => {
                    if (current()) controller.navigate(entry);
                  }}
                >
                  {words('Go to cell')}
                </button>
              </li>
            ))}
          </ol>
        </>
      )}
    </section>
  );
}
