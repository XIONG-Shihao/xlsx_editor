import { useEffect, useRef, useState } from 'react';
import type {
  WorkbookPrintCommand,
  WorkbookPrintDefinition,
  WorkbookPrintSettings,
  WorkbookPrintLayout,
  WorkbookPrintPage,
  WorkbookGeometryProfile,
} from '../core/contracts';
import { defaultPrintSettings, validatePrintSettings } from '../core/print-model.mjs';
import { emptyPrintHeaderFooter } from '../core/print-header-footer.mjs';
import { printDocumentHtml } from '../core/print-layout.mjs';
import { isPrintPreviewRetired, type AcceptedPrintPreview } from './useAcceptedPrintPreview';
export type PrintFacet = 'margins' | 'orientation' | 'paper' | 'area' | 'titles' | 'scale';
export interface PrintFocusRequest {
  facet: PrintFacet;
  serial: number;
  documentId: string;
  revision: string;
  sheet: string;
  guard: () => boolean;
}
export interface WorkbookPrintJob extends Omit<WorkbookPrintLayout, 'pages'> {
  id: string;
  pageCount: number;
  pages: WorkbookPrintPage[];
}
export interface WorkbookPrintProps {
  documentId: string;
  revision: string;
  sheet: string;
  selection: string;
  readOnly?: boolean;
  busy?: boolean;
  canPrint?: boolean;
  canExportPdf?: boolean;
  definition: (sheet: string) => Promise<WorkbookPrintDefinition>;
  onBatch: (commands: WorkbookPrintCommand[]) => Promise<unknown>;
  preview: (
    settings: WorkbookPrintSettings,
    requestId: string,
    geometryProfile?: WorkbookGeometryProfile
  ) => Promise<WorkbookPrintJob>;
  page: (jobId: string, number: number) => Promise<WorkbookPrintPage>;
  publish: (jobId: string, kind: 'pdf' | 'printer', requestId: string) => Promise<unknown>;
  release?: (jobId: string) => Promise<unknown>;
  cancel?: (requestId: string) => Promise<unknown>;
  previewSession?: AcceptedPrintPreview;
  focusRequest?: PrintFocusRequest | null;
  words?: (text: string) => string;
}
/** Shared settings and bounded physical preview; hosts own snapshot/job/dialog
 * services. The component never reads a viewport or decides a PDF filename. */
export function WorkbookPrint(props: WorkbookPrintProps): JSX.Element {
  const words = props.words ?? ((text: string) => text),
    scope = `${props.documentId}\0${props.revision}\0${props.sheet}`,
    active = useRef(props),
    lifetime = useRef(scope),
    generation = useRef(0),
    jobRef = useRef<{ id: string; release?: () => Promise<unknown> } | null>(null),
    request = useRef<string | null>(null),
    currentFrame = useRef('');
  active.current = props;
  lifetime.current = scope;
  const [definition, setDefinition] = useState<WorkbookPrintDefinition | null>(null),
    [settings, setSettings] = useState<WorkbookPrintSettings>(() => ({
      ...defaultPrintSettings(),
      area: props.selection,
    })),
    [localJob, setJob] = useState<WorkbookPrintJob | null>(null),
    [localPage, setPage] = useState<WorkbookPrintPage | null>(null),
    [localNumber, setNumber] = useState(1),
    [loading, setLoading] = useState(false),
    [pending, setPending] = useState(false),
    [error, setError] = useState(''),
    [status, setStatus] = useState(''),
    [repeatRowText, setRepeatRowText] = useState(''),
    [repeatColumnText, setRepeatColumnText] = useState(''),
    [frameLoading, setFrameLoading] = useState(false);
  const job = props.previewSession ? props.previewSession.job : localJob,
    page = props.previewSession ? props.previewSession.page : localPage,
    number = props.previewSession ? props.previewSession.number : localNumber;
  const settingsRoot = useRef<HTMLFieldSetElement>(null),
    settingsNotice = useRef<HTMLParagraphElement>(null),
    handledFocus = useRef(-1);
  useEffect(() => {
    const request = props.focusRequest;
    if (!request || request.serial === handledFocus.current || loading || !definition) return;
    handledFocus.current = request.serial;
    if (
      request.documentId !== props.documentId ||
      request.revision !== props.revision ||
      request.sheet !== props.sheet ||
      !request.guard()
    )
      return;
    const node = settingsRoot.current?.querySelector<HTMLElement>(`[data-print-facet="${request.facet}"]`);
    if (node && !node.matches(':disabled')) node.focus({ preventScroll: true });
    else settingsNotice.current?.focus({ preventScroll: true });
  }, [props.focusRequest, loading, definition]);
  const previewContainer = useRef<HTMLDivElement>(null),
    [previewWidth, setPreviewWidth] = useState(320);
  useEffect(() => {
    const container = previewContainer.current;
    if (!container) return;
    const observer = new ResizeObserver(([entry]) => {
      if (entry.contentRect.width > 0) setPreviewWidth(entry.contentRect.width);
    });
    observer.observe(container);
    return () => observer.disconnect();
  }, [job?.id]);
  const release = () => {
    if (active.current.previewSession) {
      active.current.previewSession.release(job?.id);
      return;
    }
    const old = jobRef.current;
    jobRef.current = null;
    if (old?.release) void old.release().catch(() => {});
  };
  useEffect(() => {
    const token = ++generation.current;
    setDefinition(null);
    setJob(null);
    setPage(null);
    setFrameLoading(false);
    setPending(false);
    setError('');
    setStatus('');
    request.current = null;
    if (!active.current.previewSession) release();
    setLoading(true);
    const callbacks = active.current;
    callbacks
      .definition(props.sheet)
      .then((value) => {
        if (token === generation.current) {
          setDefinition(value);
          setSettings(
            value.settings?.area
              ? value.settings
              : { ...(value.settings ?? defaultPrintSettings()), area: callbacks.selection }
          );
          setRepeatRowText(
            value.settings?.repeatRows ? `${value.settings.repeatRows.first}:${value.settings.repeatRows.last}` : ''
          );
          setRepeatColumnText(
            value.settings?.repeatColumns
              ? `${value.settings.repeatColumns.first}:${value.settings.repeatColumns.last}`
              : ''
          );
        }
      })
      .catch((failure) => {
        if (token === generation.current) setError(String(failure.message ?? failure));
      })
      .finally(() => {
        if (token === generation.current) setLoading(false);
      });
    return () => {
      generation.current++;
      lifetime.current = '';
      if (!callbacks.previewSession) release();
      request.current = null;
    };
  }, [scope]);
  const change = (value: WorkbookPrintSettings) => {
    if (scope !== lifetime.current) return;
    setSettings(value);
    setJob(null);
    setPage(null);
    setStatus('');
    release();
  };
  const draft = (): WorkbookPrintSettings => {
    const row = repeatRowText.split(':'),
      column = repeatColumnText.toUpperCase().split(':');
    if (row.length > 2 || column.length > 2) throw Error(words('Use one contiguous row/column range.'));
    return validatePrintSettings({
      ...settings,
      repeatRows: repeatRowText ? { first: Number(row[0]), last: Number(row[1] ?? row[0]) } : null,
      repeatColumns: repeatColumnText ? { first: column[0], last: column[1] ?? column[0] } : null,
    });
  };
  const publishOutput = (kind: 'pdf' | 'printer', requestId: string): Promise<unknown> => {
    if (!job || JSON.stringify(draft()) !== JSON.stringify(validatePrintSettings(job.settings)))
      return Promise.reject(Error(words('Preview pages to use these settings')));
    return active.current.previewSession
      ? active.current.previewSession.publish(job.id, kind, requestId)
      : active.current.publish(job.id, kind, requestId);
  };
  const run = async (action: (requestId: string) => Promise<unknown>, success: string, track = false) => {
    if (scope !== lifetime.current) return;
    const token = lifetime.current,
      serial = ++generation.current,
      requestId = crypto.randomUUID();
    if (track) request.current = requestId;
    setPending(true);
    setError('');
    setStatus('');
    try {
      const result = await action(requestId);
      if (token === lifetime.current && serial === generation.current)
        setStatus(
          result && typeof result === 'object' && 'cancelled' in result && result.cancelled
            ? words('Print output cancelled.')
            : success
        );
    } catch (failure) {
      if (!isPrintPreviewRetired(failure) && token === lifetime.current && serial === generation.current)
        setError(String((failure as Error).message ?? failure));
    } finally {
      if (token === lifetime.current && serial === generation.current) {
        request.current = null;
        setPending(false);
      }
    }
  };
  const preview = async () => {
    if (scope !== lifetime.current) return;
    const callbacks = active.current,
      token = lifetime.current,
      serial = ++generation.current,
      requestId = crypto.randomUUID();
    request.current = requestId;
    setPending(true);
    setError('');
    setStatus('');
    if (!callbacks.previewSession) {
      release();
      setJob(null);
      setPage(null);
    }
    try {
      const value = await (callbacks.previewSession
        ? callbacks.previewSession.prepare(draft(), requestId, { origin: { kind: 'draft', range: draft().area! } })
        : callbacks.preview(draft(), requestId));
      if (token !== lifetime.current || serial !== generation.current) {
        if (!callbacks.previewSession && callbacks.release) void callbacks.release(value.id).catch(() => {});
        return;
      }
      if (!callbacks.previewSession)
        jobRef.current = { id: value.id, release: callbacks.release ? () => callbacks.release!(value.id) : undefined };
      if (!callbacks.previewSession) {
        setJob(value);
        setPage(value.pages.find((page) => page.number === 1) ?? null);
        setNumber(1);
      }
      setFrameLoading(true);
      setStatus(words('Preview uses the accepted workbook revision.'));
    } catch (failure) {
      if (!isPrintPreviewRetired(failure) && token === lifetime.current && serial === generation.current)
        setError(String((failure as Error).message ?? failure));
    } finally {
      if (token === lifetime.current && serial === generation.current) {
        request.current = null;
        setPending(false);
      }
    }
  };
  const choosePage = async (next: number) => {
    if (!job || next < 1 || next > job.pageCount) return;
    const token = lifetime.current,
      serial = ++generation.current,
      callbacks = active.current;
    setLoading(true);
    setError('');
    try {
      if (callbacks.previewSession) {
        await callbacks.previewSession.choosePage(next);
        if (token === lifetime.current && serial === generation.current) setFrameLoading(true);
        return;
      }
      const value = job.pages.find((page) => page.number === next) ?? (await callbacks.page(job.id, next));
      if (value.number !== next) throw Error(words('The print page response did not match the requested page.'));
      if (token === lifetime.current && serial === generation.current) {
        setFrameLoading(true);
        setNumber(next);
        setPage(value);
      }
    } catch (failure) {
      if (!isPrintPreviewRetired(failure) && token === lifetime.current && serial === generation.current)
        setError(String((failure as Error).message ?? failure));
    } finally {
      if (token === lifetime.current && serial === generation.current) setLoading(false);
    }
  };
  const cancel = async () => {
    const requestId = request.current,
      token = lifetime.current,
      serial = generation.current,
      callbacks = active.current;
    if (!requestId || !callbacks.cancel) return;
    if (callbacks.previewSession?.requestId === requestId) {
      await callbacks.previewSession.cancelPreparation();
      return;
    }
    try {
      const result = await callbacks.cancel(requestId);
      if (token !== lifetime.current || serial !== generation.current || request.current !== requestId) return;
      const reply =
        result && typeof result === 'object'
          ? (result as { granted?: boolean; reason?: string; outcome?: string })
          : undefined;
      if (
        reply?.granted === false ||
        reply?.reason === 'finishing' ||
        reply?.reason === 'completed' ||
        reply?.outcome === 'finishing' ||
        reply?.outcome === 'completed'
      ) {
        setStatus(words('Print output is finishing…'));
        return;
      }
      generation.current++;
      request.current = null;
      setPending(false);
      setStatus(words('Print preparation cancelled.'));
    } catch (failure) {
      if (!isPrintPreviewRetired(failure) && token === lifetime.current && serial === generation.current)
        setError(String((failure as Error).message ?? failure));
    }
  };
  let draftMatchesJob = false;
  try {
    draftMatchesJob = !!job && JSON.stringify(draft()) === JSON.stringify(validatePrintSettings(job.settings));
  } catch {
    /* Invalid draft needs a new valid preview. */
  }
  const disabled =
      !!props.busy ||
      pending ||
      loading ||
      frameLoading ||
      !!props.previewSession?.preparing ||
      !!props.previewSession?.pageLoading ||
      !definition,
    protectedSettings = !!definition && !definition.canEdit;
  const frameKey = job && page ? `${job.id}:${page.number}` : '';
  currentFrame.current = frameKey;
  const previewHtml = job && page ? printDocumentHtml({ ...job, pages: [page] }, { pages: [page.number] }) : '';
  const pageWidth = job ? (job.pageSizePt[0] * 96) / 72 : 1,
    pageHeight = job ? (job.pageSizePt[1] * 96) / 72 : 1;
  const previewScale = Math.min(1, previewWidth / pageWidth);
  return (
    <section aria-label={words('Print and PDF')} style={{ padding: 12, display: 'grid', gap: 10 }}>
      <p ref={settingsNotice} tabIndex={-1}>
        {words('One worksheet, up to 10,000 source cells and 100 pages. Hidden rows and columns are excluded.')}
      </p>
      {definition?.protectedReasons.map((reason, index) => (
        <p role="alert" key={index}>
          {reason}
        </p>
      ))}
      <fieldset
        ref={settingsRoot}
        disabled={disabled || protectedSettings}
        style={{ display: 'flex', flexWrap: 'wrap', gap: 10 }}
      >
        <legend>{words('Print settings')}</legend>
        <label>
          {words('Print area')}{' '}
          <input
            data-print-facet="area"
            value={settings.area ?? ''}
            onChange={(event) => change({ ...settings, area: event.target.value })}
          />
        </label>
        <button
          type="button"
          onClick={() => change({ ...settings, area: props.selection, repeatRows: null, repeatColumns: null })}
        >
          {words('Use selection for print area')}
        </button>
        <label>
          {words('Paper')}{' '}
          <select
            data-print-facet="paper"
            aria-label={words('Paper')}
            value={settings.paper}
            onChange={(event) => change({ ...settings, paper: event.target.value as 'A4' | 'Letter' })}
          >
            <option>A4</option>
            <option>Letter</option>
          </select>
        </label>
        <label>
          {words('Orientation')}{' '}
          <select
            data-print-facet="orientation"
            aria-label={words('Orientation')}
            value={settings.orientation}
            onChange={(event) => change({ ...settings, orientation: event.target.value as 'portrait' | 'landscape' })}
          >
            <option value="portrait">{words('Portrait')}</option>
            <option value="landscape">{words('Landscape')}</option>
          </select>
        </label>
        {(['top', 'right', 'bottom', 'left', 'header', 'footer'] as const).map((key) => (
          <label key={key}>
            {words(key + ' margin (pt)')}{' '}
            <input
              data-print-facet={key === 'top' ? 'margins' : undefined}
              type="number"
              min={0}
              max={144}
              value={settings.marginsPt[key]}
              onChange={(event) =>
                change({ ...settings, marginsPt: { ...settings.marginsPt, [key]: Number(event.target.value) } })
              }
            />
          </label>
        ))}
        <label>
          {words('Scaling')}{' '}
          <select
            data-print-facet="scale"
            aria-label={words('Scaling')}
            value={settings.scaling.mode}
            onChange={(event) =>
              change({
                ...settings,
                scaling:
                  event.target.value === 'fit' ? { mode: 'fit', width: 1, height: 0 } : { mode: 'scale', percent: 100 },
              })
            }
          >
            <option value="scale">{words('Percent')}</option>
            <option value="fit">{words('Fit pages')}</option>
          </select>
        </label>
        {settings.scaling.mode === 'scale' ? (
          <label>
            {words('Scale (%)')}{' '}
            <input
              type="number"
              min={10}
              max={400}
              value={settings.scaling.percent}
              onChange={(event) =>
                change({ ...settings, scaling: { mode: 'scale', percent: Number(event.target.value) } })
              }
            />
          </label>
        ) : (
          <>
            <label>
              {words('Pages wide')}{' '}
              <input
                type="number"
                min={1}
                max={10}
                value={settings.scaling.width}
                onChange={(event) =>
                  change({
                    ...settings,
                    scaling: {
                      mode: 'fit',
                      width: Number(event.target.value),
                      height: settings.scaling.mode === 'fit' ? settings.scaling.height : 0,
                    },
                  })
                }
              />
            </label>
            <label>
              {words('Pages tall (0 = any)')}{' '}
              <input
                type="number"
                min={0}
                max={10}
                value={settings.scaling.height}
                onChange={(event) =>
                  change({
                    ...settings,
                    scaling: {
                      mode: 'fit',
                      width: settings.scaling.mode === 'fit' ? settings.scaling.width : 1,
                      height: Number(event.target.value),
                    },
                  })
                }
              />
            </label>
          </>
        )}
        <label>
          {words('Repeat rows')}{' '}
          <input
            data-print-facet="titles"
            placeholder="1:2"
            value={repeatRowText}
            onChange={(event) => {
              setRepeatRowText(event.target.value);
              setJob(null);
              setPage(null);
              release();
            }}
          />
        </label>
        <label>
          {words('Repeat columns')}{' '}
          <input
            placeholder="A:B"
            value={repeatColumnText}
            onChange={(event) => {
              setRepeatColumnText(event.target.value);
              setJob(null);
              setPage(null);
              release();
            }}
          />
        </label>
        <label>
          {words('Header/footer preset')}{' '}
          <select
            aria-label={words('Header/footer preset')}
            value=""
            onChange={(event) => {
              const hf = emptyPrintHeaderFooter();
              if (event.target.value === 'title') hf.header.center = props.sheet;
              else if (event.target.value === 'confidential') {
                hf.header.center = props.sheet;
                hf.footer.center = 'Confidential';
              } else if (event.target.value === 'draft') {
                hf.header.center = 'Draft';
                hf.footer.right = props.sheet;
              }
              change({ ...settings, headerFooter: hf });
            }}
          >
            <option value="">{words('Choose preset…')}</option>
            <option value="title">{words('Worksheet title')}</option>
            <option value="confidential">{words('Confidential')}</option>
            <option value="draft">{words('Draft')}</option>
            <option value="none">{words('None')}</option>
          </select>
        </label>
        {(['header', 'footer'] as const).map((side) => (
          <fieldset key={side} style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
            <legend>{words(side === 'header' ? 'Header' : 'Footer')}</legend>
            {(['left', 'center', 'right'] as const).map((align) => (
              <label key={align}>
                {words(side + ' ' + align)}{' '}
                <input
                  aria-label={words(side + ' ' + align)}
                  maxLength={255}
                  value={settings.headerFooter?.[side][align] ?? ''}
                  onChange={(event) => {
                    const hf = settings.headerFooter ?? emptyPrintHeaderFooter();
                    change({
                      ...settings,
                      headerFooter: { ...hf, [side]: { ...hf[side], [align]: event.target.value } },
                    });
                  }}
                />
              </label>
            ))}
          </fieldset>
        ))}
        <label>
          <input
            type="checkbox"
            checked={settings.headerFooter?.scaleWithDoc ?? true}
            onChange={(event) =>
              change({
                ...settings,
                headerFooter: {
                  ...(settings.headerFooter ?? emptyPrintHeaderFooter()),
                  scaleWithDoc: event.target.checked,
                },
              })
            }
          />
          {words('Scale header/footer with worksheet')}
        </label>
        <p>
          {words(
            'Single-line plain text, up to 255 encoded characters per header/footer. Formatting, images and dynamic fields stay protected.'
          )}
        </p>
        <label>
          <input
            type="checkbox"
            checked={settings.gridlines}
            onChange={(event) => change({ ...settings, gridlines: event.target.checked })}
          />
          {words('Print gridlines')}
        </label>
        <button
          type="button"
          disabled={!!props.readOnly}
          onClick={() =>
            void run(
              () =>
                active.current.onBatch([
                  {
                    operation: 'setPrintSettings',
                    sheet: props.sheet,
                    expectedPrintRevision: definition!.revision,
                    settings: draft(),
                  },
                ]),
              words('Print settings applied.')
            )
          }
        >
          {words('Apply print settings')}
        </button>
        <button
          type="button"
          disabled={!!props.readOnly || !definition?.settings?.area}
          onClick={() =>
            void run(
              () =>
                active.current.onBatch([
                  {
                    operation: 'setPrintSettings',
                    sheet: props.sheet,
                    expectedPrintRevision: definition!.revision,
                    settings: { ...definition!.settings!, area: null, repeatRows: null, repeatColumns: null },
                  },
                ]),
              words('Print area and repeats cleared.')
            )
          }
        >
          {words('Clear print area and repeats')}
        </button>
        <button type="button" onClick={() => void preview()}>
          {words('Preview pages')}
        </button>
      </fieldset>
      {pending && <p role="status">{words('Preparing print output…')}</p>}
      {(loading || frameLoading) && <p role="status">{words('Loading print page…')}</p>}
      {(error || props.previewSession?.error) && <p role="alert">{error || props.previewSession?.error}</p>}
      {(status || props.previewSession?.status) && <p role="status">{status || props.previewSession?.status}</p>}
      {pending && request.current && props.cancel && (
        <button type="button" onClick={() => void cancel()}>
          {words('Cancel print preparation')}
        </button>
      )}
      {job && !draftMatchesJob && <p role="status">{words('Preview pages to use these settings')}</p>}
      {job && (
        <div>
          <span>
            {words('Page')} {number} / {job.pageCount}
          </span>{' '}
          <button
            type="button"
            disabled={loading || frameLoading || number <= 1}
            onClick={() => void choosePage(number - 1)}
          >
            {words('Previous page')}
          </button>{' '}
          <button
            type="button"
            disabled={loading || frameLoading || number >= job.pageCount}
            onClick={() => void choosePage(number + 1)}
          >
            {words('Next page')}
          </button>{' '}
          <button
            type="button"
            disabled={disabled || !draftMatchesJob || props.canExportPdf === false}
            title={props.canExportPdf === false ? words('PDF export requires permission to create a file.') : undefined}
            onClick={() =>
              void run((requestId) => publishOutput('pdf', requestId), words('PDF output finished.'), true)
            }
          >
            {words('Export PDF')}
          </button>{' '}
          <button
            type="button"
            disabled={disabled || !draftMatchesJob || !props.canPrint}
            title={!props.canPrint ? words('Native printer dispatch is unavailable here.') : undefined}
            onClick={() =>
              void run((requestId) => publishOutput('printer', requestId), words('Print job submitted.'), true)
            }
          >
            {words('Print')}
          </button>
        </div>
      )}
      {job && page && (
        <div ref={previewContainer} className="sheet-print-preview" style={{ height: pageHeight * previewScale }}>
          <iframe
            key={frameKey}
            onLoad={(event) => {
              const marker = event.currentTarget.contentDocument?.querySelector('[data-xlsx-print-page]');
              if (
                lifetime.current === scope &&
                currentFrame.current === frameKey &&
                marker?.getAttribute('data-xlsx-print-page') === String(page.number)
              )
                setFrameLoading(false);
            }}
            title={words('Physical print page') + ' ' + number}
            sandbox="allow-same-origin"
            srcDoc={previewHtml}
            style={{
              width: pageWidth,
              height: pageHeight,
              position: 'absolute',
              top: 0,
              left: 0,
              transform: `scale(${previewScale})`,
              transformOrigin: 'top left',
              border: 0,
              background: '#fff',
            }}
          />
        </div>
      )}
    </section>
  );
}
