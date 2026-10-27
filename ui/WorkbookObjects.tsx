import { CHART_STYLES } from '../core/chart-styles.mjs';
import { WorkbookChartRecommendations } from './WorkbookChartRecommendations';
import { WorkbookInkPreview } from './WorkbookInkPreview';
import type { ChartRecommendationFrame } from './chartRecommendationsController';
import { chartSeriesAppearance } from '../core/chart-styles.mjs';
import { WorkbookModel3DPreview } from './WorkbookModel3DPreview';
import { useEffect, useRef, useState } from 'react';
import { WorkbookSmartArtPreview } from './WorkbookSmartArtPreview';
import { WorkbookShapePreview } from './WorkbookShapePreview';
import { chartPreviewGeometry } from '../core/chart-layout.mjs';
import type {
  WorkbookChartModel,
  WorkbookObjectCommand,
  WorkbookObjectPage,
  WorkbookObjectPreview,
  WorkbookObjectSummary,
  WorkbookPngAsset,
} from '../core/contracts';

export interface WorkbookObjectsProps {
  mode?: 'manage' | 'recommendations';
  recommendationFrame?: ChartRecommendationFrame;
  onRecommendationOpen?: () => void;
  documentId: string;
  revision: string;
  ownerIdentity?: string;
  sheet: string;
  selection: string;
  readOnly?: boolean;
  busy?: boolean;
  objects: (sheet: string, options?: { after?: string; limit?: number }) => Promise<WorkbookObjectPage>;
  object: (sheet: string, id: string) => Promise<WorkbookObjectPreview>;
  onBatch: (commands: WorkbookObjectCommand[]) => Promise<unknown>;
  pickPng?: () => Promise<WorkbookPngAsset | null>;
  words?: (text: string) => string;
}
/** A bounded semantic preview. Hosts may also position this view on the worksheet
 * using the owner's one-cell geometry; no grid screenshot or file URL is used. */
export function WorkbookObjectPreviewView({
  object,
  embedded = false,
  words = (text: string) => text,
}: {
  object: WorkbookObjectPreview;
  embedded?: boolean;
  words?: (text: string) => string;
}): JSX.Element {
  if (object.stroke)
    return (
      <figure
        style={{
          margin: 0,
          width: embedded ? '100%' : object.stroke.widthEmu / 9525,
          height: embedded ? '100%' : object.stroke.heightEmu / 9525,
        }}
      >
        <WorkbookInkPreview stroke={object.stroke} label={object.altText || object.name} />
        {!embedded && (
          <figcaption>{words('Stroke — choose Draw → Select Strokes or Eraser to move or remove it.')}</figcaption>
        )}
      </figure>
    );
  if (object.asset3d)
    return (
      <figure style={{ margin: 0, width: embedded ? '100%' : undefined, height: embedded ? '100%' : 240 }}>
        <WorkbookModel3DPreview
          asset={object.asset3d}
          width={(object.anchor?.widthEmu ?? 3048000) / 9525}
          height={(object.anchor?.heightEmu ?? 2286000) / 9525}
          label={object.altText || object.name}
        />
        {!embedded && (
          <figcaption>
            {words(
              object.canEdit ? '3D model — edit color and rotation from Insert → 3D Models.' : '3D model — view only.'
            )}
          </figcaption>
        )}
      </figure>
    );
  if (object.smartart)
    return (
      <figure style={{ margin: 0, width: embedded ? '100%' : undefined, height: embedded ? '100%' : 240 }}>
        <WorkbookSmartArtPreview
          model={object.smartart}
          width={(object.anchor?.widthEmu ?? 5715000) / 9525}
          height={(object.anchor?.heightEmu ?? 3429000) / 9525}
          label={object.altText || object.name}
        />
        {!embedded && (
          <figcaption>
            {words(
              object.canEdit
                ? 'SmartArt — edit nodes and layout from Insert → SmartArt.'
                : 'SmartArt sample — view only. Editing is not available yet.'
            )}
          </figcaption>
        )}
      </figure>
    );
  if (object.protectedReason) return <p>{object.protectedReason}</p>;
  if (object.shape)
    return (
      <figure style={{ margin: 0, width: embedded ? '100%' : undefined, height: embedded ? '100%' : 160 }}>
        <WorkbookShapePreview
          shape={object.shape}
          width={(object.anchor?.widthEmu ?? 2286000) / 9525}
          height={(object.anchor?.heightEmu ?? 1143000) / 9525}
          label={object.altText || object.shape.text || object.name}
        />
        {!embedded && <figcaption>{object.name}</figcaption>}
      </figure>
    );
  if (object.image)
    return (
      <figure style={{ margin: 0, width: embedded ? '100%' : undefined, height: embedded ? '100%' : undefined }}>
        <img
          src={`data:image/png;base64,${object.image.base64}`}
          alt={object.altText || object.name}
          decoding="async"
          loading="lazy"
          style={{
            width: embedded ? '100%' : undefined,
            height: embedded ? '100%' : undefined,
            maxWidth: '100%',
            maxHeight: embedded ? undefined : 350,
            objectFit: 'contain',
          }}
        />
        {!embedded && <figcaption>{object.name}</figcaption>}
      </figure>
    );
  if (!object.data || !object.chart) return <p>{words('Preview is unavailable.')}</p>;
  const { categories, series, hidden } = object.data,
    geometry = chartPreviewGeometry(object),
    title = object.chart.title || object.name;
  const { width, height, left, plotWidth } = geometry,
    appearance = (index: number) => chartSeriesAppearance(object.chart!.style, index);
  return (
    <figure style={{ margin: 0, width: embedded ? '100%' : undefined, height: embedded ? '100%' : undefined }}>
      <svg
        role="img"
        aria-label={title}
        viewBox={`0 0 ${width} ${height}`}
        style={{
          width: '100%',
          height: embedded ? '100%' : undefined,
          maxWidth: embedded ? undefined : 600,
          display: 'block',
        }}
      >
        <title>{title}</title>
        <text x={width / 2} y={20} textAnchor="middle" fontSize={15} fill="currentColor">
          {title}
        </text>
        {geometry.lines.map(([x1, y1, x2, y2], i) => (
          <line key={i} x1={x1} y1={y1} x2={x2} y2={y2} stroke="currentColor" opacity={0.45} />
        ))}
        {geometry.bars.map((bar, i) => (
          <rect
            key={i}
            x={bar.x}
            y={bar.y}
            width={bar.width}
            height={bar.height}
            fill={appearance(bar.series).fill}
            stroke={appearance(bar.series).stroke}
            strokeWidth={appearance(bar.series).width}
            strokeDasharray={appearance(bar.series).dash}
          >
            <title>{bar.title}</title>
          </rect>
        ))}
        {geometry.segments.map((segment, i) =>
          segment.points.length === 1 ? (
            <circle
              key={i}
              cx={segment.points[0][0]}
              cy={segment.points[0][1]}
              r={2.5}
              fill={appearance(segment.series).color}
            />
          ) : (
            <polyline
              key={i}
              points={segment.points.map((point) => point.join(',')).join(' ')}
              fill="none"
              stroke={appearance(segment.series).color}
              strokeWidth={appearance(segment.series).width ?? 2}
              strokeDasharray={appearance(segment.series).dash}
            />
          )
        )}
        {geometry.labels.map((label, i) => (
          <text key={i} x={label.x} y={label.y} textAnchor={label.anchor} fontSize={11} fill="currentColor">
            {label.text}
          </text>
        ))}
        {object.chart.legend !== 'none' &&
          series.map((item, index) => (
            <g key={index} transform={`translate(${left + index * (plotWidth / series.length)},${height - 12})`}>
              <rect
                width={9}
                height={9}
                y={-8}
                fill={appearance(index).fill}
                stroke={appearance(index).stroke}
                strokeWidth={appearance(index).width}
                strokeDasharray={appearance(index).dash}
              />
              <text x={13} fontSize={10} fill="currentColor">
                {item.name.slice(0, 30)}
              </text>
            </g>
          ))}
      </svg>
      {!embedded && (
        <details>
          <summary>{words('Chart data')}</summary>
          <table>
            <caption>{title}</caption>
            <thead>
              <tr>
                <th scope="col">{words('Category')}</th>
                {series.map((item, index) => (
                  <th key={index} scope="col">
                    {item.name}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {categories.map((category, index) => (
                <tr key={index}>
                  <th scope="row">
                    {category}
                    {hidden[index] ? ` (${words('hidden')})` : ''}
                  </th>
                  {series.map((item, column) => (
                    <td key={column}>
                      {item.values[index] === null
                        ? ''
                        : Object.is(item.values[index], -0)
                          ? '-0'
                          : String(item.values[index])}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </details>
      )}
    </figure>
  );
}

/** Shared create/update/delete controls. All writes are owner commands; every
 * asynchronous response is fenced by document+revision+sheet lifetime. */
export function WorkbookObjects(props: WorkbookObjectsProps): JSX.Element {
  const words = props.words ?? ((text: string) => text),
    scope = JSON.stringify([
      props.documentId,
      props.ownerIdentity ?? null,
      props.revision,
      props.sheet,
      props.mode ?? 'manage',
    ]),
    generation = useRef(0),
    choice = useRef(0),
    lifetime = useRef(scope),
    snapshotScope = useRef(scope),
    writing = useRef(false),
    callbacks = useRef(props);
  callbacks.current = props;
  lifetime.current = scope;
  const [storedPage, setPage] = useState<WorkbookObjectPage>({ objects: [], total: 0, nextCursor: null }),
    [cursors, setCursors] = useState<Array<string | undefined>>([undefined]),
    [refresh, setRefresh] = useState(0),
    [storedSelected, setSelected] = useState<WorkbookObjectSummary | null>(null),
    [storedPreview, setPreview] = useState<WorkbookObjectPreview | null>(null),
    [storedError, setError] = useState(''),
    [storedLoading, setLoading] = useState(false),
    [storedPending, setPending] = useState(false),
    [storedMode, setMode] = useState<'create' | 'edit' | 'image' | null>(null),
    [model, setModel] = useState<WorkbookChartModel>({
      kind: 'column',
      range: props.selection,
      title: '',
      legend: 'bottom',
    }),
    [cell, setCell] = useState('E2'),
    [imageWidth, setImageWidth] = useState(600),
    [altText, setAltText] = useState('');
  const fresh = snapshotScope.current === scope,
    page = fresh ? storedPage : { objects: [], total: 0, nextCursor: null },
    selected = fresh ? storedSelected : null,
    preview = fresh ? storedPreview : null,
    mode = fresh ? storedMode : null,
    error = fresh ? storedError : '',
    loading = !fresh || storedLoading,
    pending = storedPending;
  // A commit under a new head/session must never paint a previous chart while
  // passive cleanup is still pending. The state reset remains in the effect.
  useEffect(() => {
    snapshotScope.current = scope;
    lifetime.current = scope;
    choice.current++;
    setPage({ objects: [], total: 0, nextCursor: null });
    setCursors([undefined]);
    setSelected(null);
    setPreview(null);
    setMode(null);
    setError('');
    return () => {
      lifetime.current = '';
      choice.current++;
      generation.current++;
    };
  }, [scope]);
  const after = fresh ? cursors.at(-1) : undefined;
  useEffect(() => {
    const token = ++generation.current;
    setLoading(true);
    setError('');
    callbacks.current
      .objects(props.sheet, { after, limit: 50 })
      .then((value) => {
        if (token === generation.current) setPage(value);
      })
      .catch((failure) => {
        if (token === generation.current) setError(String(failure.message ?? failure));
      })
      .finally(() => {
        if (token === generation.current) setLoading(false);
      });
    return () => {
      generation.current++;
    };
  }, [scope, after, refresh]);
  const write = async (action: () => Promise<unknown>): Promise<void> => {
    if (lifetime.current !== scope) throw Error('The original chart workbook changed.');
    if (writing.current) throw Error('Wait for the chart operation to finish.');
    writing.current = true;
    setPending(true);
    try {
      await action();
    } finally {
      writing.current = false;
      setPending(false);
    }
  };
  const run = async (action: () => Promise<unknown>): Promise<void> => {
    if (lifetime.current !== scope) return;
    const token = scope;
    setError('');
    try {
      await write(action);
      if (token === lifetime.current) {
        setMode(null);
        setRefresh((value) => value + 1);
      }
    } catch (failure) {
      if (token === lifetime.current) setError(String((failure as Error).message ?? failure));
    }
  };
  const choose = async (entry: WorkbookObjectSummary): Promise<void> => {
    if (lifetime.current !== scope) return;
    const token = ++choice.current,
      active = scope;
    setSelected(entry);
    setPreview(null);
    setMode(null);
    setError('');
    try {
      const value = await callbacks.current.object(props.sheet, entry.id);
      if (token === choice.current && active === lifetime.current) setPreview(value);
    } catch (failure) {
      if (token === choice.current && active === lifetime.current)
        setError(String((failure as Error).message ?? failure));
    }
  };
  const disabled = !!props.readOnly || !!props.busy || pending;
  return (
    <section
      className="sheet-objects"
      aria-label={words('Workbook objects')}
      style={{ padding: 12, display: 'grid', gap: 10 }}
    >
      {props.mode === 'recommendations' && props.recommendationFrame && (
        <WorkbookChartRecommendations
          frame={props.recommendationFrame}
          pending={pending || !!props.busy}
          words={words}
          onAccept={async (frame, command) => {
            const current = callbacks.current.recommendationFrame;
            if (
              !current ||
              current.key !== frame.key ||
              !current.active ||
              current.reason ||
              props.readOnly ||
              props.busy
            )
              throw Error(current?.reason ?? 'The original chart preview changed.');
            await write(() => frame.apply(structuredClone(command)));
          }}
        />
      )}
      <div hidden={props.mode === 'recommendations'}>
        <label>
          {words('Image anchor cell')}{' '}
          <input
            value={cell}
            disabled={disabled}
            onChange={(event) => setCell(event.target.value)}
            aria-label={words('Image anchor cell')}
          />
        </label>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button
            type="button"
            disabled={disabled}
            onClick={() => {
              setMode('create');
              setModel({ kind: 'column', range: props.selection, title: '', legend: 'bottom', style: 'colorful' });
              setCell('E2');
            }}
          >
            {words('Create chart')}
          </button>
          {props.onRecommendationOpen && (
            <button type="button" disabled={pending} onClick={props.onRecommendationOpen}>
              {words('Recommended Charts')}
            </button>
          )}
          <button
            type="button"
            disabled={disabled || !props.pickPng}
            title={!props.pickPng ? words('Image import is unavailable here.') : undefined}
            onClick={() =>
              void run(async () => {
                const active = callbacks.current,
                  token = lifetime.current,
                  asset = await active.pickPng!();
                if (asset && token === lifetime.current)
                  await active.onBatch([
                    {
                      operation: 'insertImage',
                      sheet: active.sheet,
                      objectId: crypto.randomUUID(),
                      expectedObjectRevision: null,
                      asset,
                      anchor: { cell },
                      altText: words('Inserted image'),
                    },
                  ]);
              })
            }
          >
            {words('Insert PNG')}
          </button>
          <button type="button" disabled={loading || pending} onClick={() => setRefresh((value) => value + 1)}>
            {words('Refresh objects')}
          </button>
        </div>
        {mode && mode !== 'image' && (
          <form
            onSubmit={(event) => {
              event.preventDefault();
              void run(() =>
                callbacks.current.onBatch([
                  mode === 'edit' && selected
                    ? {
                        operation: 'updateChart',
                        sheet: props.sheet,
                        objectId: selected.id,
                        expectedObjectRevision: selected.revision,
                        patch: model,
                        anchor: { cell },
                      }
                    : {
                        operation: 'createChart',
                        sheet: props.sheet,
                        objectId: crypto.randomUUID(),
                        expectedObjectRevision: null,
                        chart: model,
                        anchor: { cell },
                      },
                ])
              );
            }}
          >
            <fieldset disabled={disabled} style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
              <legend>{words(mode === 'edit' ? 'Edit chart' : 'Create chart')}</legend>
              <label>
                {words('Chart type')}{' '}
                <select
                  aria-label={words('Chart type')}
                  value={model.kind}
                  onChange={(event) => setModel({ ...model, kind: event.target.value as 'column' | 'bar' | 'line' })}
                >
                  <option value="column">{words('Column')}</option>
                  <option value="bar">{words('Horizontal bar')}</option>
                  <option value="line">{words('Line')}</option>
                </select>
              </label>
              <label>
                {words('Chart style')}{' '}
                <select
                  aria-label={words('Chart style')}
                  value={model.style ?? ''}
                  onChange={(event) =>
                    setModel({ ...model, style: event.target.value as 'colorful' | 'monochrome' | 'outline' })
                  }
                >
                  {model.style === undefined && (
                    <option value="" disabled>
                      {words('Current default')}
                    </option>
                  )}
                  {CHART_STYLES.map((style) => (
                    <option key={style.id} value={style.id}>
                      {words(style.label)}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {words('Data range')}{' '}
                <input
                  value={model.range}
                  onChange={(event) => setModel({ ...model, range: event.target.value })}
                  required
                />
              </label>
              <label>
                {words('Chart title')}{' '}
                <input
                  value={model.title}
                  maxLength={128}
                  onChange={(event) => setModel({ ...model, title: event.target.value })}
                />
              </label>
              <label>
                {words('Legend')}{' '}
                <select
                  aria-label={words('Legend')}
                  value={model.legend ?? 'bottom'}
                  onChange={(event) => setModel({ ...model, legend: event.target.value as 'bottom' | 'none' })}
                >
                  <option value="bottom">{words('Bottom')}</option>
                  <option value="none">{words('None')}</option>
                </select>
              </label>
              <label>
                {words('Anchor cell')} <input value={cell} onChange={(event) => setCell(event.target.value)} required />
              </label>
              <p>{words('First row: series names. First column: category labels. Up to 1,000 source cells.')}</p>
              <button type="submit">{words(mode === 'edit' ? 'Apply chart changes' : 'Create chart')}</button>
              <button type="button" onClick={() => setMode(null)}>
                {words('Cancel')}
              </button>
            </fieldset>
          </form>
        )}
        {mode === 'image' && selected && preview?.image && (
          <form
            onSubmit={(event) => {
              event.preventDefault();
              const widthEmu = Math.round(imageWidth * 9525),
                heightEmu = Math.round((widthEmu * preview.image!.heightPx) / preview.image!.widthPx);
              void run(() =>
                callbacks.current.onBatch([
                  {
                    operation: 'updateImage',
                    sheet: props.sheet,
                    objectId: selected.id,
                    expectedObjectRevision: selected.revision,
                    anchor: { cell, widthEmu, heightEmu },
                    altText,
                  },
                ])
              );
            }}
          >
            <fieldset disabled={disabled} style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
              <legend>{words('Edit image placement')}</legend>
              <label>
                {words('Anchor cell')}{' '}
                <input
                  aria-label={words('Image anchor')}
                  value={cell}
                  required
                  onChange={(event) => setCell(event.target.value)}
                />
              </label>
              <label>
                {words('Width in pixels')}{' '}
                <input
                  aria-label={words('Image width')}
                  type="number"
                  min="1"
                  max="5249"
                  step="1"
                  value={imageWidth}
                  required
                  onChange={(event) => setImageWidth(Number(event.target.value))}
                />
              </label>
              <label>
                {words('Alternative text')}{' '}
                <input
                  aria-label={words('Image alternative text')}
                  maxLength={256}
                  value={altText}
                  onChange={(event) => setAltText(event.target.value)}
                />
              </label>
              <p>{words('Height follows the original image proportions.')}</p>
              <button type="submit">{words('Apply image changes')}</button>
              <button type="button" onClick={() => setMode(null)}>
                {words('Cancel')}
              </button>
            </fieldset>
          </form>
        )}
        {error && <p role="alert">{error}</p>}
        {loading && <p role="status">{words('Loading objects…')}</p>}
        {pending && <p role="status">{words('Applying object change…')}</p>}
        <ul style={{ listStyle: 'none', margin: 0, padding: 0 }}>
          {page.objects.map((entry) => (
            <li key={entry.id}>
              <button
                type="button"
                data-workbook-object={entry.id}
                aria-pressed={selected?.id === entry.id}
                onClick={() => void choose(entry)}
              >
                {entry.name} ({words(entry.kind)})
              </button>
              {entry.protectedReason && <span> — {entry.protectedReason}</span>}
            </li>
          ))}
        </ul>
        <div>
          <span>
            {page.total} {words('objects')}
          </span>{' '}
          <button
            type="button"
            disabled={loading || cursors.length === 1}
            onClick={() => setCursors((value) => value.slice(0, -1))}
          >
            {words('Previous objects')}
          </button>{' '}
          <button
            type="button"
            disabled={loading || !page.nextCursor}
            onClick={() => setCursors((value) => [...value, page.nextCursor!])}
          >
            {words('Next objects')}
          </button>
        </div>
        {selected && (
          <div>
            <strong>{selected.name}</strong>
            <p>{selected.anchor?.cell}</p>
            {selected.kind === 'chart' && (
              <button
                type="button"
                disabled={disabled || !selected.canEdit}
                onClick={() => {
                  setMode('edit');
                  setModel(selected.chart!);
                  setCell(selected.anchor?.cell ?? 'E2');
                }}
              >
                {words('Edit chart')}
              </button>
            )}
            {selected.kind === 'image' && (
              <button
                type="button"
                disabled={disabled || !selected.canEdit || !preview?.image}
                onClick={() => {
                  setMode('image');
                  setCell(selected.anchor?.cell ?? 'E2');
                  setImageWidth(Math.max(1, Math.round((selected.anchor?.widthEmu ?? 5715000) / 9525)));
                  setAltText(selected.altText ?? '');
                }}
              >
                {words('Edit image placement')}
              </button>
            )}{' '}
            <button
              type="button"
              disabled={disabled || !selected.canDelete}
              onClick={() =>
                void run(() =>
                  callbacks.current.onBatch([
                    {
                      operation: 'deleteObject',
                      sheet: props.sheet,
                      objectId: selected.id,
                      expectedObjectRevision: selected.revision,
                    },
                  ])
                )
              }
            >
              {words('Delete object')}
            </button>
          </div>
        )}
        {preview && <WorkbookObjectPreviewView object={preview} words={words} />}
      </div>
    </section>
  );
}
