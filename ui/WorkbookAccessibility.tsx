import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import { ACCESSIBILITY_LIMITS } from '../core/accessibility.mjs';
import { createAccessibilityController, type AccessibilityFrame } from './accessibilityController';
const labels = {
  'missing-alt-text': 'Missing alternative text',
  'blank-table-header': 'Blank table header',
  'duplicate-table-header': 'Duplicate table headers',
};
const PAGE = 50;
export function WorkbookAccessibility({
  frame,
  words = (s: string) => s,
}: {
  frame: AccessibilityFrame | null;
  words?: (s: string) => string;
}) {
  const [controller] = useState(createAccessibilityController),
    lifetime = useRef(0);
  const snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot);
  useLayoutEffect(() => controller.setFrame(frame));
  useEffect(() => {
    const serial = ++lifetime.current;
    return () => {
      controller.setFrame(null);
      queueMicrotask(() => {
        if (lifetime.current === serial) controller.dispose();
      });
    };
  }, [controller]);
  const report = snapshot.key === frame?.key ? snapshot.report : null;
  const [paging, setPaging] = useState({ report: null as typeof report, offset: 0, skipOffset: 0 });
  const offsets = paging.report === report ? paging : { report, offset: 0, skipOffset: 0 };
  const disabled = !frame?.active || snapshot.loading;
  const [error, setError] = useState({ key: '', text: '' });
  const run = (action: () => void) => {
    try {
      setError({ key: '', text: '' });
      action();
    } catch (e) {
      setError({ key: frame?.key ?? '', text: String(e instanceof Error ? e.message : e) });
    }
  };
  return (
    <section
      className="sheet-accessibility"
      aria-label={words('Workbook accessibility')}
      style={{ display: 'grid', gap: 8, padding: 12 }}
    >
      <p>
        {words(
          'Read-only checks for missing alternative text and blank or repeated table headers across this workbook.'
        )}
      </p>
      <button
        type="button"
        disabled={disabled}
        onClick={() => {
          if (frame) run(() => controller.refresh(frame));
        }}
      >
        {words('Check again')}
      </button>
      {snapshot.loading && <p role="status">{words('Checking workbook content…')}</p>}
      {report && (
        <>
          <h3>{words('Coverage')}</h3>
          <p role="status">
            {words(
              report.status === 'partial'
                ? 'Some content was not checked.'
                : 'All content in this check’s supported profile was checked.'
            )}
          </p>
          <p>
            {words('Objects checked')}: {report.coverage.objects.checked} /{' '}
            {report.coverage.objects.total ?? words('unknown')}; {words('skipped')}: {report.coverage.objects.skipped}
          </p>
          <p>
            {words('Tables checked')}: {report.coverage.tables.checked} /{' '}
            {report.coverage.tables.total ?? words('unknown')}; {words('skipped')}: {report.coverage.tables.skipped}
          </p>
          <h3>
            {words('Findings')} ({report.findings.length})
          </h3>
          {!report.findings.length ? (
            <p>
              {words(
                report.status === 'partial'
                  ? 'No findings in checked content; skipped content still needs review.'
                  : 'No findings in checked content.'
              )}
            </p>
          ) : (
            <>
              <ol start={offsets.offset + 1}>
                {report.findings.slice(offsets.offset, offsets.offset + PAGE).map((f, i) => {
                  const index = offsets.offset + i,
                    hidden = !frame?.visibleSheets.includes(f.sheet);
                  return (
                    <li key={index} style={{ marginBottom: 14 }}>
                      <strong>{words(labels[f.rule])}</strong>
                      <p>{words(f.message)}</p>
                      {f.tableName && (
                        <p>
                          {words('Table')}: {f.tableName}
                        </p>
                      )}
                      <button
                        type="button"
                        aria-label={words('Go to finding') + ' ' + f.sheet + '!' + f.cell}
                        disabled={disabled || hidden}
                        onClick={() => {
                          if (frame) run(() => controller.navigate(frame, index));
                        }}
                      >
                        {f.sheet + '!' + f.cell}
                      </button>
                      {f.relatedCells
                        .filter((cell) => cell !== f.cell)
                        .map((cell) => (
                          <button
                            key={cell}
                            type="button"
                            aria-label={words('Go to related header') + ' ' + f.sheet + '!' + cell}
                            disabled={disabled || hidden}
                            onClick={() => {
                              if (frame) run(() => controller.navigate(frame, index, cell));
                            }}
                          >
                            {cell}
                          </button>
                        ))}
                      {hidden && (
                        <p>{words('This finding is on a hidden worksheet. Unhide it separately to navigate.')}</p>
                      )}
                    </li>
                  );
                })}
              </ol>
              <div>
                <button
                  type="button"
                  disabled={disabled || !offsets.offset}
                  onClick={() => setPaging({ ...offsets, offset: Math.max(0, offsets.offset - PAGE) })}
                >
                  {words('Previous findings')}
                </button>{' '}
                {offsets.offset + 1}–{Math.min(report.findings.length, offsets.offset + PAGE)} /{' '}
                {report.findings.length}{' '}
                <button
                  type="button"
                  disabled={disabled || offsets.offset + PAGE >= report.findings.length}
                  onClick={() => setPaging({ ...offsets, offset: offsets.offset + PAGE })}
                >
                  {words('More findings')}
                </button>
              </div>
            </>
          )}
          <details>
            <summary>{words('Coverage details and limits')}</summary>
            <p>
              {words('Header cell positions read')}: {report.coverage.headerCellsRead}
            </p>
            <p>
              {words('Limits')}: {ACCESSIBILITY_LIMITS.tables} {words('tables')}, {ACCESSIBILITY_LIMITS.columns}{' '}
              {words('columns per table')}, {ACCESSIBILITY_LIMITS.headerCells} {words('header cell positions')}.{' '}
              {words('Imported/extended objects and formula or special table headers may be skipped.')}
            </p>
            <p>{words('These content checks do not certify full accessibility. Repairs are separate actions.')}</p>
          </details>
          {!!report.coverage.skipped.length && (
            <details>
              <summary>
                {words('Skipped content')} ({report.coverage.skipped.length})
              </summary>
              <ul>
                {report.coverage.skipped.slice(offsets.skipOffset, offsets.skipOffset + PAGE).map((s, i) => (
                  <li key={offsets.skipOffset + i}>
                    {[s.sheet, s.name].filter(Boolean).join(' — ') ||
                      words(s.kind === 'table' ? 'Table inventory' : 'Object inventory')}
                    : {words(s.reason)}
                  </li>
                ))}
              </ul>
              <button
                type="button"
                disabled={disabled || !offsets.skipOffset}
                onClick={() => setPaging({ ...offsets, skipOffset: Math.max(0, offsets.skipOffset - PAGE) })}
              >
                {words('Previous skipped items')}
              </button>{' '}
              <button
                type="button"
                disabled={disabled || offsets.skipOffset + PAGE >= report.coverage.skipped.length}
                onClick={() => setPaging({ ...offsets, skipOffset: offsets.skipOffset + PAGE })}
              >
                {words('More skipped items')}
              </button>
            </details>
          )}
        </>
      )}
      {((snapshot.key === frame?.key && snapshot.error) || (error.key === frame?.key && error.text)) && (
        <p role="alert">{words((snapshot.key === frame?.key && snapshot.error) || error.text)}</p>
      )}
    </section>
  );
}
