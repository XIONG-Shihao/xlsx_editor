/** One accepted print-job lease shared by the pane and page mode. */
import { useLayoutEffect, useRef, useState } from 'react';
import type { WorkbookPrintSettings, WorkbookPrintPage } from '../core/contracts';
import type { WorkbookPrintJob } from './WorkbookPrint';
import { printDocumentHtml } from '../core/print-layout.mjs';
import { printPaperSize, validatePrintSettings } from '../core/print-model.mjs';
import { validateGeometryProfile } from '../core/worksheet-geometry.mjs';

export interface PrintPreviewScope {
  documentId: string;
  grantDocumentId: string;
  state: string;
  sheet: string;
  ownerKey: string;
  activationEpoch: number;
  active: boolean;
  geometryKey?: string;
}
export type AcceptedPrintJob = WorkbookPrintJob & { state: string };
export interface PrintPreviewCallbacks {
  preview: (settings: WorkbookPrintSettings, requestId: string) => Promise<WorkbookPrintJob>;
  page: (jobId: string, number: number) => Promise<WorkbookPrintPage>;
  publish?: (jobId: string, kind: 'pdf' | 'printer', requestId: string) => Promise<unknown>;
  release?: (jobId: string) => Promise<unknown>;
  cancel?: (requestId: string) => Promise<unknown>;
}
export interface PrintAreaOrigin {
  kind: 'saved' | 'selection' | 'draft';
  range: string;
}
export interface PreparePageOptions {
  guard?: () => boolean;
  origin?: PrintAreaOrigin;
}
export interface AcceptedPrintPreview {
  job: AcceptedPrintJob | null;
  page: WorkbookPrintPage | null;
  number: number;
  preparing: boolean;
  pageLoading: boolean;
  publishing: boolean;
  requestId: string | null;
  status: string;
  error: string;
  origin: PrintAreaOrigin | null;
  prepare: (
    settings: WorkbookPrintSettings,
    requestId?: string,
    options?: PreparePageOptions
  ) => Promise<AcceptedPrintJob>;
  reuseCurrent: (settings: WorkbookPrintSettings, options?: PreparePageOptions) => AcceptedPrintJob | null;
  choosePage: (number: number) => Promise<void>;
  publish: (jobId: string, kind: 'pdf' | 'printer', requestId: string) => Promise<unknown>;
  cancelPreparation: () => Promise<void>;
  abandonPreparation: () => void;
  release: (expectedJobId?: string) => void;
}
export class PrintPreviewRetired extends Error {
  readonly code = 'PRINT_PREVIEW_RETIRED';
  constructor() {
    super('This page preparation was retired.');
    this.name = 'PrintPreviewRetired';
  }
}
export function isPrintPreviewRetired(error: unknown): boolean {
  return error instanceof PrintPreviewRetired;
}
interface Lease {
  key: string;
  job: AcceptedPrintJob;
  callbacks: PrintPreviewCallbacks;
  released: boolean;
  publication?: Promise<unknown>;
}
interface Preparation {
  key: string;
  serial: number;
  requestId: string;
  callbacks: PrintPreviewCallbacks;
}
const keyFor = (scope: PrintPreviewScope) =>
  JSON.stringify([
    scope.documentId,
    scope.grantDocumentId,
    scope.state,
    scope.sheet,
    scope.ownerKey,
    scope.activationEpoch,
    scope.active,
    scope.geometryKey ?? null,
  ]);
const message = (error: unknown) => (error instanceof Error ? error.message : String(error));
function renderable(job: AcceptedPrintJob, page: WorkbookPrintPage, number: number): void {
  const box = (value: unknown): boolean =>
    Array.isArray(value) &&
    value.length === 4 &&
    value.every((item) => typeof item === 'number' && Number.isFinite(item)) &&
    value[2] > 0 &&
    value[3] > 0;
  if (
    !page ||
    page.number !== number ||
    !Array.isArray(page.rowIndices) ||
    !Array.isArray(page.columnIndices) ||
    !Array.isArray(page.cells) ||
    !Array.isArray(page.objects) ||
    page.rowIndices.some((index) => !Number.isInteger(index) || index < 0 || index >= 1048576) ||
    page.columnIndices.some((index) => !Number.isInteger(index) || index < 0 || index >= 16384) ||
    new Set(page.rowIndices).size !== page.rowIndices.length ||
    new Set(page.columnIndices).size !== page.columnIndices.length ||
    page.cells.length > 10000 ||
    page.objects.length > 32 ||
    page.cells.some((cell) => typeof cell.text !== 'string' || typeof cell.ref !== 'string' || !box(cell.bbox)) ||
    page.objects.some((object) => !box(object.bbox) || !box(object.clipBox))
  )
    throw Error('The accepted page contains invalid geometry or content.');
  printDocumentHtml({ ...job, pages: [page] }, { pages: [number] });
}
function admitted(
  value: WorkbookPrintJob,
  scope: PrintPreviewScope,
  requested: WorkbookPrintSettings
): AcceptedPrintJob {
  const job = value as AcceptedPrintJob;
  if (
    !job ||
    typeof job.id !== 'string' ||
    !job.id ||
    job.state !== scope.state ||
    job.sheet !== scope.sheet ||
    !Number.isInteger(job.pageCount) ||
    job.pageCount < 1 ||
    job.pageCount > 100 ||
    !Array.isArray(job.pages) ||
    !job.pages.some((page) => page.number === 1) ||
    job.units !== 'pt' ||
    !Array.isArray(job.pageSizePt) ||
    job.pageSizePt.length !== 2 ||
    job.pageSizePt.some((size) => typeof size !== 'number' || !Number.isFinite(size) || size <= 0 || size > 2000) ||
    typeof job.scale !== 'number' ||
    !Number.isFinite(job.scale) ||
    job.scale < 0.1 ||
    job.scale > 4 ||
    !/^[0-9a-f]{64}$/i.test(job.sourceFingerprint)
  )
    throw Error('The accepted print job did not match this workbook revision.');
  if (scope.geometryKey) {
    const expected = validateGeometryProfile(JSON.parse(scope.geometryKey));
    if (!job.geometryProfile) throw Error('The print preview did not confirm the worksheet dimensions.');
    const actual = validateGeometryProfile(job.geometryProfile, expected.font);
    if (JSON.stringify(actual) !== JSON.stringify(expected))
      throw Error('The accepted pages changed the requested worksheet dimensions.');
  }
  const settings = validatePrintSettings(job.settings),
    size = printPaperSize(requested);
  if (
    JSON.stringify(settings) !== JSON.stringify(requested) ||
    job.pageSizePt.some((value, index) => Math.abs(value - size[index]) > 0.000001) ||
    (requested.scaling.mode === 'scale' && job.scale !== requested.scaling.percent / 100)
  )
    throw Error('The accepted pages changed the requested print settings.');
  const accepted = { ...job, pages: [job.pages.find((page) => page.number === 1)!] };
  renderable(accepted, accepted.pages[0], 1);
  return accepted;
}
async function retire(lease: Lease): Promise<void> {
  if (lease.released) return;
  lease.released = true;
  await lease.publication?.catch(() => {});
  // A failed cleanup cannot turn an accepted preview into a workbook failure.
  await lease.callbacks.release?.(lease.job.id).catch(() => {});
}
export function useAcceptedPrintPreview({
  scope,
  callbacks,
}: {
  scope: PrintPreviewScope;
  callbacks: PrintPreviewCallbacks | null;
}): AcceptedPrintPreview {
  const key = keyFor(scope),
    live = useRef({ key, scope, callbacks });
  live.current = { key, scope, callbacks };
  const feedbackKey = useRef(key),
    generation = useRef(0),
    lease = useRef<Lease | null>(null),
    preparation = useRef<Preparation | null>(null),
    pageSerial = useRef(0),
    publication = useRef<Lease | null>(null);
  const [job, setJob] = useState<AcceptedPrintJob | null>(null),
    [page, setPage] = useState<WorkbookPrintPage | null>(null),
    [number, setNumber] = useState(1);
  const [publishing, setPublishing] = useState(false),
    [preparing, setPreparing] = useState(false),
    [pageLoading, setPageLoading] = useState(false),
    [requestId, setRequestId] = useState<string | null>(null);
  const [status, setStatus] = useState(''),
    [error, setError] = useState(''),
    [origin, setOrigin] = useState<PrintAreaOrigin | null>(null);
  const current = (capturedKey: string, serial?: number) =>
    live.current.scope.active &&
    live.current.key === capturedKey &&
    (serial === undefined || generation.current === serial);
  const stop = (cancel = true): void => {
    if (live.current.key !== key) return;
    generation.current++;
    pageSerial.current++;
    const active = preparation.current;
    preparation.current = null;
    if (active && cancel) void active.callbacks.cancel?.(active.requestId).catch(() => {});
    setPreparing(false);
    setPageLoading(false);
    setRequestId(null);
  };
  const release = (expectedJobId?: string): void => {
    if (live.current.key !== key) return;
    const previous = lease.current;
    if ((previous && previous.key !== key) || (expectedJobId && previous?.job.id !== expectedJobId)) return;
    stop();
    lease.current = null;
    if (previous) void retire(previous);
    setJob(null);
    setPage(null);
    setOrigin(null);
    setStatus('');
  };
  useLayoutEffect(() => {
    release();
    setError('');
    setNumber(1);
    return () => {
      const active = preparation.current;
      if (active?.key === key) {
        generation.current++;
        preparation.current = null;
        void active.callbacks.cancel?.(active.requestId).catch(() => {});
      }
      const previous = lease.current;
      if (previous?.key === key) {
        pageSerial.current++;
        lease.current = null;
        void retire(previous);
      }
    };
  }, [key]);
  const reuseCurrent = (settings: WorkbookPrintSettings, options: PreparePageOptions = {}): AcceptedPrintJob | null => {
    const captured = live.current,
      previous = lease.current;
    if (
      captured.key !== key ||
      !captured.scope.active ||
      !captured.callbacks ||
      !previous ||
      previous.key !== key ||
      previous.released ||
      preparation.current ||
      publication.current ||
      previous.job.state !== captured.scope.state ||
      previous.job.sheet !== captured.scope.sheet ||
      (options.guard && !options.guard())
    )
      return null;
    const requested = validatePrintSettings(settings);
    if (JSON.stringify(validatePrintSettings(previous.job.settings)) !== JSON.stringify(requested)) return null;
    // The same immutable pages may now be reached through a saved area or a
    // captured selection. Change only that UI provenance, never owner settings.
    feedbackKey.current = key;
    setOrigin(options.origin ?? { kind: 'draft', range: requested.area! });
    setError('');
    setStatus('Preview uses the accepted workbook revision.');
    return previous.job;
  };
  const prepare = async (
    settings: WorkbookPrintSettings,
    id: string = crypto.randomUUID(),
    options: PreparePageOptions = {}
  ): Promise<AcceptedPrintJob> => {
    const captured = live.current;
    if (captured.key !== key) throw new PrintPreviewRetired();
    if (!captured.scope.active || !captured.callbacks)
      throw Error('Page Layout preview is unavailable in this editor.');
    if (preparation.current || publication.current) throw Error('Wait for the current print operation.');
    if (options.guard && !options.guard()) throw new PrintPreviewRetired();
    const requested = validatePrintSettings(settings),
      previous = lease.current;
    lease.current = null;
    if (previous) void retire(previous);
    feedbackKey.current = captured.key;
    setJob(null);
    setPage(null);
    setError('');
    setStatus('Preparing accepted pages…');
    setOrigin(null);
    const serial = ++generation.current,
      active = { key: captured.key, serial, requestId: id, callbacks: captured.callbacks };
    preparation.current = active;
    setPreparing(true);
    setRequestId(id);
    let returned: WorkbookPrintJob | undefined;
    try {
      returned = await captured.callbacks.preview(requested, id);
      if (!current(captured.key, serial) || (options.guard && !options.guard())) throw new PrintPreviewRetired();
      const accepted = admitted(returned, captured.scope, requested);
      lease.current = { key: captured.key, job: accepted, callbacks: captured.callbacks, released: false };
      returned = undefined;
      setJob(accepted);
      setPage(accepted.pages[0]);
      setNumber(1);
      setOrigin(options.origin ?? { kind: 'draft', range: requested.area! });
      setStatus('Preview uses the accepted workbook revision.');
      return accepted;
    } catch (failure) {
      if (returned && typeof returned.id === 'string') await captured.callbacks.release?.(returned.id).catch(() => {});
      if (!current(captured.key, serial) || (options.guard && !options.guard())) throw new PrintPreviewRetired();
      if (!(failure instanceof PrintPreviewRetired)) setError(message(failure));
      throw failure;
    } finally {
      if (preparation.current === active) {
        preparation.current = null;
        if (current(captured.key, serial)) {
          setPreparing(false);
          setRequestId(null);
        }
      }
    }
  };
  const choosePage = async (next: number): Promise<void> => {
    const captured = lease.current;
    if (
      !captured ||
      captured.key !== key ||
      !current(captured.key) ||
      !Number.isInteger(next) ||
      next < 1 ||
      next > captured.job.pageCount
    )
      throw Error('This print page is unavailable or expired.');
    const serial = ++pageSerial.current;
    feedbackKey.current = captured.key;
    setPageLoading(true);
    setError('');
    try {
      const value =
        captured.job.pages.find((page) => page.number === next) ??
        (await captured.callbacks.page(captured.job.id, next));
      if (lease.current !== captured || !current(captured.key) || pageSerial.current !== serial)
        throw new PrintPreviewRetired();
      renderable(captured.job, value, next);
      setPage(value);
      setNumber(next);
    } catch (failure) {
      if (lease.current === captured && current(captured.key) && pageSerial.current === serial)
        setError(message(failure));
      throw failure;
    } finally {
      if (lease.current === captured && current(captured.key) && pageSerial.current === serial) setPageLoading(false);
    }
  };
  const publish = async (id: string, kind: 'pdf' | 'printer', request: string): Promise<unknown> => {
    const captured = lease.current;
    if (!captured || captured.key !== key || !current(key) || captured.job.id !== id || !captured.callbacks.publish)
      throw Error('This accepted print job is unavailable or expired.');
    if (publication.current || preparation.current) throw Error('Wait for the current print operation.');
    publication.current = captured;
    setPublishing(true);
    const promise = Promise.resolve().then(() => captured.callbacks.publish!(id, kind, request));
    captured.publication = promise;
    try {
      return await promise;
    } finally {
      captured.publication = undefined;
      if (publication.current === captured) {
        publication.current = null;
        setPublishing(false);
      }
    }
  };
  const cancelPreparation = async (): Promise<void> => {
    const active = preparation.current;
    if (!active || active.key !== key || live.current.key !== key || !active.callbacks.cancel) return;
    try {
      const reply = await active.callbacks.cancel(active.requestId);
      if (preparation.current !== active || !current(active.key, active.serial)) return;
      const result =
        reply && typeof reply === 'object'
          ? (reply as { granted?: boolean; cancelled?: boolean; reason?: string; outcome?: string })
          : {};
      if (result.granted === false || ['finishing', 'completed'].includes(result.reason ?? result.outcome ?? '')) {
        setStatus('Print preparation is finishing…');
        return;
      }
      if (result.granted !== true && result.cancelled !== true && result.outcome !== 'cancelled') {
        setStatus('Cancellation was not confirmed; waiting for print preparation.');
        return;
      }
      stop(false);
      setStatus('Print preparation cancelled. Accepted workbook edits were kept.');
    } catch (failure) {
      if (preparation.current === active && current(active.key, active.serial)) setError(message(failure));
      throw failure;
    }
  };
  // Render can precede scope cleanup. Never expose an old owner's page/status
  // for that first render, including hidden inactive tab trees.
  const visible = scope.active && lease.current?.key === key && lease.current.job.id === job?.id;
  const activePreparation = scope.active && preparation.current?.key === key;
  const feedback = scope.active && feedbackKey.current === key;
  return {
    job: visible ? job : null,
    page: visible ? page : null,
    number: visible ? number : 1,
    preparing: !!activePreparation && preparing,
    publishing: !!visible && publication.current === lease.current && publishing,
    pageLoading: !!visible && pageLoading,
    requestId: activePreparation ? requestId : null,
    status: feedback ? status : '',
    error: feedback ? error : '',
    origin: visible ? origin : null,
    prepare,
    reuseCurrent,
    choosePage,
    publish,
    cancelPreparation,
    abandonPreparation: () => {
      if (live.current.key !== key) return;
      stop();
      setStatus('Returned to Normal view.');
      setError('');
    },
    release,
  };
}
