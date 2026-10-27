import { useEffect, useRef, useState } from 'react';
import type {
  LocalQuery,
  LocalQueryConnection,
  LocalQueryDraft,
  LocalQueryList,
  LocalQueryPreview,
  LocalQueryPreviewOptions,
  LocalQueryRefreshOptions,
} from '../core/local-query.mjs';
import type { QueryRefreshAllControls } from './queryRefreshAllController';
import type { TextConversionPolicy } from '../core/text-conversion.mjs';
export type QueryPreview = LocalQueryPreview;
export interface WorkbookQueryControls {
  documentId: string;
  revision: string;
  ownerId?: string;
  refreshAll?: QueryRefreshAllControls;
  list: () => Promise<LocalQueryList>;
  create: (draft: LocalQueryDraft) => Promise<{ connection?: LocalQueryConnection; cancelled?: boolean }>;
  update: (
    id: string,
    expectedRevision: number,
    draft: LocalQueryDraft
  ) => Promise<{ connection: LocalQueryConnection }>;
  remove: (id: string, expectedRevision: number) => Promise<{ removed: boolean }>;
  preview: (
    id: string,
    expectedRevision: number,
    requestId: string,
    options?: LocalQueryPreviewOptions
  ) => Promise<{ queryPreview: QueryPreview }>;
  refresh: (
    id: string,
    expectedRevision: number,
    requestId: string,
    options?: LocalQueryRefreshOptions
  ) => Promise<{
    refreshed?: boolean;
    cancelled?: boolean;
    warning?: string;
    outputRows?: number;
    outputColumns?: number;
  }>;
  cancel: (requestId: string) => Promise<{ outcome?: string; granted?: boolean; reason?: string }>;
}
export interface WorkbookQueriesProps extends WorkbookQueryControls {
  sheet: string;
  selection: string;
  readOnly?: boolean;
  busy?: boolean;
  visible?: boolean;
}
type Filter = LocalQuery['filters'][number];
const blankFilter = (): Filter => ({ column: 1, op: 'equals', value: '', caseSensitive: true });
const typeLabels: Record<TextConversionPolicy, string> = {
  text: 'Text',
  number: 'Number',
  boolean: 'Boolean',
  'iso-date': 'ISO Date',
};
export function WorkbookQueries(props: WorkbookQueriesProps): JSX.Element {
  const [list, setList] = useState<LocalQueryList>({ entries: [] }),
    [selected, setSelected] = useState(''),
    [name, setName] = useState('Local query'),
    [start, setStart] = useState(props.selection.split(':')[0]),
    [sheet, setSheet] = useState(props.sheet),
    [columns, setColumns] = useState(''),
    [filters, setFilters] = useState<Filter[]>([]),
    [header, setHeader] = useState(true),
    [policies, setPolicies] = useState<TextConversionPolicy[]>([]),
    [fieldCount, setFieldCount] = useState(0),
    [fieldPage, setFieldPage] = useState(0),
    [preview, setPreview] = useState<{ binding: string; settings: string; data: QueryPreview }>(),
    [message, setMessage] = useState(''),
    [running, setRunning] = useState<string>(),
    [finishing, setFinishing] = useState(false);
  const owner = props.documentId + '\0' + (props.ownerId ?? ''),
    binding = owner + '\0' + props.revision + '\0' + (props.visible !== false),
    epoch = useRef(0),
    live = useRef(binding),
    mounted = useRef(true),
    boundOwner = useRef(owner),
    liveOwner = useRef(owner),
    listSerial = useRef(0),
    request = useRef<{ id: string; owner: string; cancel: WorkbookQueryControls['cancel'] } | null>(null);
  live.current = binding;
  liveOwner.current = owner;
  const disabled = Boolean(props.visible === false || props.readOnly || props.busy || running),
    chosen = list.entries.find((r) => r.id === selected);
  function draft(): LocalQueryDraft {
    const selectedColumns = columns.trim() ? columns.split(',').map((v) => Number(v.trim())) : [];
    if (
      selectedColumns.some((v) => !Number.isInteger(v) || v < 1) ||
      new Set(selectedColumns).size !== selectedColumns.length
    )
      throw Error('Use unique positive column numbers separated by commas.');
    return {
      name,
      target: { sheet, start },
      query: { columns: selectedColumns, filters, keepHeader: header, ...(policies.length ? { policies } : {}) },
    };
  }
  let settings = '';
  try {
    settings = JSON.stringify(draft());
  } catch {
    /* Invalid settings cannot be previewed or refreshed. */
  }
  const form = JSON.stringify([selected, settings, chosen?.revision ?? null]),
    liveForm = useRef(form);
  liveForm.current = form;
  const formCurrent = () => liveForm.current === form;
  const savedSettings = chosen ? JSON.stringify({ name: chosen.name, target: chosen.target, query: chosen.query }) : '',
    dirty = Boolean(chosen && settings !== savedSettings),
    currentPreview = preview?.binding === binding && preview.settings === settings ? preview.data : undefined,
    inspection = currentPreview?.conversionInspection,
    replacements = currentPreview?.replacementInspection,
    typed = policies.some((policy) => policy !== 'text'),
    baseRefreshReason = dirty
      ? 'Save query settings before Refresh.'
      : typed && !inspection
        ? 'Preview and review typed conversions before Refresh.'
        : inspection?.rejected
          ? 'Correct rejected conversions or choose Text before Refresh.'
          : replacements?.blocked
            ? 'Protected destination cells block Refresh.'
            : undefined,
    refreshReason =
      baseRefreshReason ??
      (replacements?.occupied ? 'Review and explicitly replace occupied destination cells.' : undefined);
  const liveReview = useRef(inspection?.revision);
  liveReview.current = inspection?.revision;
  const visibleFields = Math.min(256, columns.trim() ? columns.split(',').length : fieldCount),
    sync = async () => {
      if (!mounted.current || live.current !== binding) return;
      const serial = ++listSerial.current,
        result = await props.list();
      if (mounted.current && live.current === binding && serial === listSerial.current) setList(result);
    };
  const edit = (change: () => void) => {
    if (disabled || live.current !== binding || !formCurrent()) return;
    epoch.current++;
    setPreview(undefined);
    setMessage('');
    change();
  };
  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      live.current = '';
      epoch.current++;
      listSerial.current++;
      request.current = null;
    };
  }, []);
  useEffect(() => {
    epoch.current++;
    setPreview(undefined);
    if (boundOwner.current !== owner) {
      boundOwner.current = owner;
      request.current = null;
      setRunning(undefined);
      setFinishing(false);
      setList({ entries: [] });
      setSelected('');
      setName('Local query');
      setSheet(props.sheet);
      setStart(props.selection.split(':')[0]);
      setColumns('');
      setFilters([]);
      setHeader(true);
      setPolicies([]);
      setFieldCount(0);
      setFieldPage(0);
      setMessage('');
    }
    if (props.visible !== false)
      void sync().catch((e) => {
        if (mounted.current && live.current === binding) setMessage(String(e instanceof Error ? e.message : e));
      });
  }, [owner, props.revision, props.visible]);
  function choose(r: LocalQueryConnection): void {
    epoch.current++;
    setSelected(r.id);
    setName(r.name);
    setSheet(r.target.sheet);
    setStart(r.target.start);
    setColumns(r.query.columns.join(','));
    setFilters(r.query.filters.map((f) => ({ ...f })));
    setHeader(r.query.keepHeader);
    setPolicies(r.query.policies ?? []);
    setFieldCount(
      r.id === selected
        ? Math.max(fieldCount, r.query.columns.length || r.last?.outputColumns || r.query.policies?.length || 0)
        : r.query.columns.length || r.last?.outputColumns || r.query.policies?.length || 0
    );
    setFieldPage(0);
    setPreview(undefined);
    setMessage('');
  }
  async function action(
    kind: 'create' | 'update' | 'remove' | 'preview' | 'refresh',
    offset?: number,
    replacementOffset?: number,
    replace = false
  ): Promise<void> {
    if (disabled || live.current !== binding || !formCurrent() || request.current) return;
    if (
      (kind === 'refresh' || offset !== undefined || replacementOffset !== undefined) &&
      liveReview.current !== inspection?.revision
    )
      return;
    if ((kind === 'preview' || kind === 'refresh') && dirty) {
      setMessage('Save query settings before previewing or refreshing.');
      return;
    }
    if (kind === 'refresh' && (baseRefreshReason || (replacements?.occupied && !replace))) return;
    const token = ++epoch.current,
      scope = binding,
      requestId = crypto.randomUUID(),
      owned = { id: requestId, owner, cancel: props.cancel };
    request.current = owned;
    setRunning(requestId);
    setFinishing(false);
    setMessage('');
    const current = () => mounted.current && request.current === owned && liveOwner.current === owned.owner,
      currentRevision = () => current() && token === epoch.current && live.current === scope;
    try {
      if (kind === 'create') {
        const result = await props.create(draft());
        if (!currentRevision()) return;
        if (result.connection) choose(result.connection);
        else setMessage('Source chooser cancelled.');
      } else {
        if (!chosen) throw Error('Select a saved query.');
        if (kind === 'update') {
          const result = await props.update(chosen.id, chosen.revision, draft());
          if (currentRevision()) choose(result.connection);
        } else if (kind === 'remove') {
          await props.remove(chosen.id, chosen.revision);
          if (currentRevision()) {
            setSelected('');
            setPreview(undefined);
          }
        } else if (kind === 'preview') {
          const options: LocalQueryPreviewOptions =
            offset !== undefined
              ? { offset, expectedConversionRevision: inspection?.revision }
              : replacementOffset !== undefined
                ? { replacementOffset, expectedConversionRevision: replacements?.revision }
                : {};
          if ((offset !== undefined || replacementOffset !== undefined) && !inspection)
            throw Error('Preview the query again before paging its conversion issues.');
          const result = await props.preview(chosen.id, chosen.revision, requestId, options);
          if (currentRevision()) {
            if (result.queryPreview.connectionId !== chosen.id || result.queryPreview.revision !== chosen.revision)
              throw Error('Query preview belongs to a different saved definition.');
            setFieldCount(result.queryPreview.outputColumns);
            setPreview({ binding: scope, settings, data: result.queryPreview });
          }
        } else {
          const result = await props.refresh(chosen.id, chosen.revision, requestId, {
            ...(inspection ? { expectedConversionRevision: inspection.revision } : {}),
            ...(replace && replacements ? { expectedReplacementRevision: replacements.revision } : {}),
          });
          if (current())
            setMessage(
              result.warning ??
                (result.cancelled
                  ? 'Refresh cancelled. Accepted edits were kept.'
                  : result.refreshed
                    ? 'Refreshed ' +
                      result.outputRows +
                      ' rows and ' +
                      result.outputColumns +
                      ' columns. One Undo restores the previous result.'
                    : 'Refresh did not complete.')
            );
          if (current()) setPreview(undefined);
        }
      }
      if (kind !== 'refresh' && current() && live.current === scope) await sync();
    } catch (e) {
      if (current() && live.current === scope) setMessage(String(e instanceof Error ? e.message : e));
    } finally {
      if (current()) {
        request.current = null;
        setRunning(undefined);
        setFinishing(false);
      }
    }
  }
  async function cancel(): Promise<void> {
    const owned = request.current;
    if (!owned || finishing) return;
    const current = () => mounted.current && request.current === owned && liveOwner.current === owned.owner;
    try {
      const result = await owned.cancel(owned.id);
      if (!current()) return;
      if (result.outcome === 'finishing' || result.reason === 'finishing') {
        setFinishing(true);
        setMessage('Refresh is finishing its accepted recovery write.');
      } else if (result.outcome === 'completed' || result.reason === 'completed') {
        setFinishing(true);
        setMessage('Query output is completing. Waiting for its accepted result.');
      } else if (result.granted === false)
        setMessage('Query cancellation was not granted' + (result.reason ? ': ' + result.reason : '.'));
      else {
        epoch.current++;
        setPreview(undefined);
        setMessage('Cancellation requested. Accepted edits are kept.');
      }
    } catch (e) {
      if (current() && live.current === binding) setMessage(String(e instanceof Error ? e.message : e));
    }
  }
  return (
    <div className="workbook-queries" aria-label="Queries and refresh" data-formula-draft-zone="true">
      <p>
        Choose a local CSV or TSV. Filters compare source text. Text keeps leading zeros and formula-looking text; other
        field types convert only when explicitly selected.
      </p>
      {list.warning && <p role="alert">{list.warning}</p>}
      <label>
        Saved query{' '}
        <select
          aria-label="Saved query"
          value={selected}
          disabled={disabled}
          onChange={(e) => {
            if (live.current !== binding) return;
            const r = list.entries.find((r) => r.id === e.target.value);
            if (r) choose(r);
            else
              edit(() => {
                setSelected('');
                setPolicies([]);
                setFieldCount(0);
              });
          }}
        >
          <option value="">New connection</option>
          {list.entries.map((r) => (
            <option key={r.id} value={r.id}>
              {r.name}
              {r.pending ? ' (pending receipt)' : ''}
            </option>
          ))}
        </select>
      </label>
      {chosen && (
        <p>
          {chosen.source.name} → {chosen.target.sheet}!{chosen.target.start}
          {chosen.last ? ' · last result ' + chosen.last.outputRows + ' × ' + chosen.last.outputColumns : ''}
        </p>
      )}
      <label>
        Query name{' '}
        <input
          aria-label="Query name"
          value={name}
          disabled={disabled}
          onChange={(e) => edit(() => setName(e.target.value))}
        />
      </label>
      <label>
        Query worksheet{' '}
        <input
          aria-label="Query worksheet"
          value={sheet}
          disabled={disabled}
          onChange={(e) => edit(() => setSheet(e.target.value))}
        />
      </label>
      <label>
        Query start cell{' '}
        <input
          aria-label="Query start cell"
          value={start}
          disabled={disabled}
          onChange={(e) => edit(() => setStart(e.target.value.toUpperCase()))}
        />
      </label>
      <button
        disabled={disabled}
        onClick={() =>
          edit(() => {
            setSheet(props.sheet);
            setStart(props.selection.split(':')[0]);
          })
        }
      >
        Use selection for query
      </button>
      <label>
        Selected columns{' '}
        <input
          aria-label="Selected columns"
          placeholder="All columns, or 1,3,2"
          value={columns}
          disabled={disabled}
          onChange={(e) =>
            edit(() => {
              setColumns(e.target.value);
              setPolicies([]);
              setFieldPage(0);
            })
          }
        />
      </label>
      <label>
        <input
          aria-label="Keep first row as header"
          type="checkbox"
          checked={header}
          disabled={disabled}
          onChange={(e) => edit(() => setHeader(e.target.checked))}
        />
        Keep first row as header
      </label>
      <fieldset>
        <legend>Output field types</legend>
        <p>
          Headers and empty fields stay Text. Number requires exact decimal meaning; Boolean accepts TRUE/FALSE. ISO
          Date uses YYYY-MM-DD and the target workbook calendar.
        </p>
        {!visibleFields && <p>Preview once to discover all output fields.</p>}
        {Array.from({ length: Math.max(0, Math.min(10, visibleFields - fieldPage)) }, (_, i) => i + fieldPage).map(
          (index) => (
            <label key={index}>
              Field {index + 1}
              <select
                aria-label={'Output field ' + (index + 1) + ' type'}
                value={policies[index] ?? 'text'}
                disabled={disabled}
                onChange={(e) =>
                  edit(() =>
                    setPolicies((old) =>
                      Array.from({ length: Math.max(old.length, index + 1) }, (_, i) =>
                        i === index ? (e.target.value as TextConversionPolicy) : (old[i] ?? 'text')
                      )
                    )
                  )
                }
              >
                {Object.entries(typeLabels).map(([value, label]) => (
                  <option key={value} value={value}>
                    {label}
                  </option>
                ))}
              </select>
            </label>
          )
        )}
        {fieldPage > 0 && (
          <button disabled={disabled} onClick={() => setFieldPage((v) => Math.max(0, v - 10))}>
            Previous output fields
          </button>
        )}
        {fieldPage + 10 < visibleFields && (
          <button disabled={disabled} onClick={() => setFieldPage((v) => v + 10)}>
            More output fields
          </button>
        )}
      </fieldset>
      {filters.map((filter, index) => (
        <div key={index} aria-label={'Query filter ' + (index + 1)}>
          <label>
            Column{' '}
            <input
              aria-label={'Filter ' + (index + 1) + ' column'}
              type="number"
              min={1}
              max={16384}
              value={filter.column}
              disabled={disabled}
              onChange={(e) =>
                edit(() =>
                  setFilters((values) =>
                    values.map((v, i) => (i === index ? { ...v, column: Number(e.target.value) } : v))
                  )
                )
              }
            />
          </label>
          <select
            aria-label={'Filter ' + (index + 1) + ' operator'}
            value={filter.op}
            disabled={disabled}
            onChange={(e) =>
              edit(() =>
                setFilters((values) =>
                  values.map((v, i) => (i === index ? { ...v, op: e.target.value as Filter['op'] } : v))
                )
              )
            }
          >
            {['equals', 'notEquals', 'contains', 'isEmpty', 'notEmpty'].map((op) => (
              <option key={op} value={op}>
                {op}
              </option>
            ))}
          </select>
          <input
            aria-label={'Filter ' + (index + 1) + ' value'}
            value={filter.value}
            disabled={disabled || ['isEmpty', 'notEmpty'].includes(filter.op)}
            onChange={(e) =>
              edit(() =>
                setFilters((values) => values.map((v, i) => (i === index ? { ...v, value: e.target.value } : v)))
              )
            }
          />
          <label>
            <input
              aria-label={'Filter ' + (index + 1) + ' case sensitive'}
              type="checkbox"
              checked={filter.caseSensitive}
              disabled={disabled}
              onChange={(e) =>
                edit(() =>
                  setFilters((values) =>
                    values.map((v, i) => (i === index ? { ...v, caseSensitive: e.target.checked } : v))
                  )
                )
              }
            />
            Case sensitive
          </label>
          <button
            disabled={disabled}
            onClick={() => edit(() => setFilters((values) => values.filter((_, i) => i !== index)))}
          >
            Remove filter {index + 1}
          </button>
        </div>
      ))}
      <button
        disabled={disabled || filters.length >= 8}
        onClick={() => edit(() => setFilters((v) => [...v, blankFilter()]))}
      >
        Add query filter
      </button>
      <div>
        <button disabled={disabled} onClick={() => void action('create')}>
          Choose CSV/TSV source…
        </button>
        <button disabled={disabled || !chosen || chosen.pending} onClick={() => void action('update')}>
          Save query settings
        </button>
        <button
          title={dirty ? 'Save query settings before Preview.' : undefined}
          disabled={disabled || !chosen || chosen.pending || dirty}
          onClick={() => void action('preview')}
        >
          Preview query
        </button>
        <button
          title={refreshReason}
          disabled={disabled || !chosen || chosen.pending || Boolean(refreshReason)}
          onClick={() => void action('refresh')}
        >
          Refresh query
        </button>
        {replacements && replacements.occupied > 0 && (
          <button
            title={baseRefreshReason}
            disabled={disabled || !chosen || chosen.pending || Boolean(baseRefreshReason)}
            onClick={() => void action('refresh', undefined, undefined, true)}
          >
            Refresh and replace {replacements.occupied} occupied cells
          </button>
        )}
        <button disabled={disabled || !chosen || chosen.pending} onClick={() => void action('remove')}>
          Remove query
        </button>
        {running && (
          <button disabled={finishing} onClick={() => void cancel()}>
            Cancel query
          </button>
        )}
      </div>
      {dirty && <p role="status">Settings changed. Save query settings, then Preview again.</p>}
      {message && (
        <p role="status" aria-live="polite">
          {message}
        </p>
      )}
      {currentPreview && (
        <div aria-label="Query preview">
          <p>
            {currentPreview.outputRows} rows × {currentPreview.outputColumns} columns · {currentPreview.source.name}
            {currentPreview.truncated ? ' · bounded text preview; all included fields are inspected' : ''}
          </p>
          <table>
            <tbody>
              {currentPreview.rows.map((row, index) => (
                <tr key={index}>
                  {row.map((value, col) => (
                    <td key={col}>{value}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {inspection && (
        <div aria-label="Query conversion inspection" data-conversion-revision={inspection.revision}>
          <p>
            {inspection.rejected} rejected fields · {inspection.changed} fields with conversion notices ·{' '}
            {inspection.date1904 === null
              ? 'Date calendar unavailable'
              : 'Date calendar: ' + (inspection.date1904 ? '1904' : '1900')}
          </p>
          {inspection.rejected > 0 && (
            <p role="alert">Rejected fields block the entire Refresh. Correct the source or choose Text.</p>
          )}
          <table aria-label="Query conversion issues">
            <thead>
              <tr>
                <th>Source</th>
                <th>Destination</th>
                <th>Original text</th>
                <th>Conversion</th>
              </tr>
            </thead>
            <tbody>
              {inspection.issues.map((issue) => (
                <tr key={issue.index} aria-label={'Query conversion issue ' + issue.destinationRef}>
                  <td>{issue.sourceRef}</td>
                  <td>
                    {currentPreview?.target.sheet}!{issue.destinationRef}
                  </td>
                  <td>
                    <pre
                      style={{
                        maxWidth: '16em',
                        maxHeight: '6em',
                        overflow: 'auto',
                        whiteSpace: 'pre-wrap',
                        overflowWrap: 'anywhere',
                      }}
                    >
                      {JSON.stringify(issue.raw)}
                    </pre>
                  </td>
                  <td>
                    {typeLabels[issue.policy]} ·{' '}
                    {issue.status === 'rejected' ? issue.reason : issue.kind + ' ' + issue.display}
                    {issue.notices.map((notice, i) => (
                      <p key={i}>{notice}</p>
                    ))}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <p>
            {inspection.totalIssues ? inspection.offset + 1 : 0}–{inspection.offset + inspection.issues.length} of{' '}
            {inspection.totalIssues} issues and notices.
          </p>
          {inspection.offset > 0 && (
            <button disabled={disabled} onClick={() => void action('preview', Math.max(0, inspection.offset - 100))}>
              Previous conversion issues
            </button>
          )}
          {inspection.nextOffset !== null && (
            <button disabled={disabled} onClick={() => void action('preview', inspection.nextOffset!)}>
              Next conversion issues
            </button>
          )}
        </div>
      )}
      {replacements && (
        <div aria-label="Query destination replacements" data-replacement-revision={replacements.revision}>
          <p>
            {replacements.occupied} occupied cells in {replacements.totalNewPositions} newly claimed positions. Existing
            managed cells retain their own edit protection.
          </p>
          {replacements.blocked > 0 && (
            <p role="alert">{replacements.blocked} protected destination cells block the entire Refresh.</p>
          )}
          <table aria-label="Query replacement rows">
            <thead>
              <tr>
                <th>Source</th>
                <th>Destination</th>
                <th>Original source text</th>
                <th>Current destination</th>
                <th>Proposed output</th>
              </tr>
            </thead>
            <tbody>
              {replacements.rows.map((row) => (
                <tr key={row.index} aria-label={'Query replacement ' + row.destinationRef}>
                  <td>{row.sourceRef ?? 'Output-clearing padding'}</td>
                  <td>
                    {currentPreview?.target.sheet}!{row.destinationRef}
                  </td>
                  <td>
                    <pre
                      style={{
                        maxWidth: '16em',
                        maxHeight: '6em',
                        overflow: 'auto',
                        whiteSpace: 'pre-wrap',
                        overflowWrap: 'anywhere',
                      }}
                    >
                      {JSON.stringify(row.sourceText)}
                    </pre>
                  </td>
                  <td>
                    <pre
                      style={{
                        maxWidth: '16em',
                        maxHeight: '6em',
                        overflow: 'auto',
                        whiteSpace: 'pre-wrap',
                        overflowWrap: 'anywhere',
                      }}
                    >
                      {row.current.kind} · {row.current.display}
                    </pre>
                    {row.current.pending && <p>Pending calculation. {row.current.storedDisplay}</p>}
                    {row.current.reason && <p>{row.current.reason}</p>}
                  </td>
                  <td>
                    {row.outputKind} · {row.outputDisplay}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <p>
            {replacements.occupied ? replacements.offset + 1 : 0}–{replacements.offset + replacements.rows.length} of{' '}
            {replacements.occupied} occupied cells.
          </p>
          {replacements.offset > 0 && (
            <button
              disabled={disabled}
              onClick={() => void action('preview', undefined, Math.max(0, replacements.offset - 100))}
            >
              Previous replacements
            </button>
          )}
          {replacements.nextOffset !== null && (
            <button disabled={disabled} onClick={() => void action('preview', undefined, replacements.nextOffset!)}>
              More replacements
            </button>
          )}
        </div>
      )}
    </div>
  );
}
