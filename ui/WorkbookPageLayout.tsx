/** Actual accepted physical pages; editing remains in Normal. */
import { useRef, useState } from 'react';
import { printDocumentHtml } from '../core/print-layout.mjs';
import type { AcceptedPrintPreview } from './useAcceptedPrintPreview';
export function WorkbookPageLayout({
  preview,
  zoom = 100,
  onNormal,
  words = (value: string) => value,
}: {
  preview: AcceptedPrintPreview;
  zoom?: number;
  onNormal: () => void;
  words?: (value: string) => string;
}): JSX.Element {
  const { job, page, number } = preview,
    key = job && page ? `${job.id}:${job.state}:${page.number}` : '',
    frameKey = useRef(key);
  frameKey.current = key;
  const [loaded, setLoaded] = useState('');
  const scale = Math.max(10, Math.min(400, zoom)) / 100,
    width = ((job?.pageSizePt[0] ?? 1) * 96) / 72,
    height = ((job?.pageSizePt[1] ?? 1) * 96) / 72;
  const html = job && page ? printDocumentHtml({ ...job, pages: [page] }, { pages: [page.number] }) : '';
  return (
    <section
      className="sheet-page-layout"
      aria-label={words('Page Layout preview')}
      data-print-area-kind={preview.origin?.kind}
      data-print-area={preview.origin?.range}
    >
      <header className="sheet-page-layout-toolbar">
        <div>
          <strong>{words('Page Layout preview')}</strong>
          <p>{words('Read-only preview; edit in Normal view.')}</p>
        </div>
        <button type="button" onClick={onNormal}>
          {words('Normal')}
        </button>
        {job && (
          <>
            <span>
              {words('Page')} {number} / {job.pageCount}
            </span>
            <button
              disabled={preview.pageLoading || number <= 1}
              onClick={() => void preview.choosePage(number - 1).catch(() => {})}
            >
              {words('Previous page')}
            </button>
            <button
              disabled={preview.pageLoading || number >= job.pageCount}
              onClick={() => void preview.choosePage(number + 1).catch(() => {})}
            >
              {words('Next page')}
            </button>
          </>
        )}
        {preview.preparing && (
          <button type="button" onClick={() => void preview.cancelPreparation().catch(() => {})}>
            {words('Cancel print preparation')}
          </button>
        )}
      </header>
      {preview.origin && job && (
        <p role="status">
          {words(
            preview.origin.kind === 'saved'
              ? 'Saved print area'
              : preview.origin.kind === 'selection'
                ? 'Captured selection; print settings were not saved'
                : 'Draft print settings; not saved'
          )}{' '}
          · {job.sheet}!{preview.origin.range}
        </p>
      )}
      {preview.status && <p role="status">{words(preview.status)}</p>}
      {preview.error && <p role="alert">{preview.error}</p>}
      {!job && !preview.preparing && (
        <p>{words('No current accepted page preview. Return to Normal to prepare it again.')}</p>
      )}
      {job && page && (
        <div className="sheet-page-layout-scroll">
          <div className="sheet-page-layout-paper" style={{ width: width * scale, height: height * scale }}>
            <iframe
              key={key}
              title={words('Physical print page') + ' ' + number}
              sandbox="allow-same-origin"
              srcDoc={html}
              data-accepted-print-job={job.id}
              data-accepted-print-state={job.state}
              data-accepted-print-sheet={job.sheet}
              data-accepted-print-source={job.sourceFingerprint}
              onLoad={(event) => {
                const document = event.currentTarget.contentDocument,
                  marker = document?.querySelector('[data-xlsx-print-page]'),
                  fingerprint = document?.querySelector('meta[name="xlsx-print-source"]')?.getAttribute('content');
                if (
                  frameKey.current === key &&
                  marker?.getAttribute('data-xlsx-print-page') === String(number) &&
                  fingerprint === job.sourceFingerprint
                )
                  setLoaded(key);
              }}
              style={{
                width,
                height,
                transform: `scale(${scale})`,
                transformOrigin: 'top left',
                border: 0,
                background: '#fff',
              }}
            />
            {loaded !== key && <span role="status">{words('Loading print page…')}</span>}
          </div>
        </div>
      )}
    </section>
  );
}
