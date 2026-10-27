import { useId } from 'react';
import type { AcceptedPrintPreview } from './useAcceptedPrintPreview';
import type { PageBoundaryModel } from './pageBoundaries';
import { projectPageBoundary } from './pageBoundaries';
import type { ViewGeometry, ViewViewport } from './viewGeometry';
export function PageBoundaryToolbar({
  preview,
  model,
  onNormal,
  onPage,
  frozen,
  words = (text: string) => text,
}: {
  preview: AcceptedPrintPreview;
  model: PageBoundaryModel | null;
  onNormal: () => void;
  onPage: (number: number) => void;
  frozen: boolean;
  words?: (text: string) => string;
}) {
  return (
    <section
      className="sheet-page-boundary-toolbar"
      aria-label={words('Page Break Preview')}
      data-accepted-print-job={preview.job?.id}
      data-accepted-print-state={preview.job?.state}
    >
      <strong>{words('Page Break Preview')}</strong>
      <span>{words('Read-only boundaries; edit in Normal. Break dragging is unavailable.')}</span>
      <button type="button" onClick={onNormal}>
        {words('Normal')}
      </button>
      {model && (
        <>
          <label>
            {words('Go to page')}{' '}
            <select
              aria-label={words('Go to page boundary')}
              defaultValue=""
              onChange={(event) => {
                if (event.target.value) onPage(Number(event.target.value));
              }}
            >
              <option value="">{words('Choose a page')}</option>
              {model.pages.map((page) => (
                <option key={page.number} value={page.number}>
                  {words('Page')} {page.number} · {page.range}
                </option>
              ))}
            </select>
          </label>
          <small>
            {model.area} · {model.pages.length} {words('pages')}
            {model.repeatRows ? ' · ' + words('Repeated rows') + ' ' + model.repeatRows : ''}
            {model.repeatColumns ? ' · ' + words('Repeated columns') + ' ' + model.repeatColumns : ''}
          </small>
        </>
      )}
      {frozen && <small>{words('Frozen panes are suspended in this map; Normal retains them.')}</small>}
      {!model && (
        <span role="status">{words('No current accepted boundaries. Return to Normal to prepare them again.')}</span>
      )}
    </section>
  );
}
export function PageBoundaryOverlay({
  model,
  geometry,
  viewport,
  words = (text: string) => text,
}: {
  model: PageBoundaryModel;
  geometry: ViewGeometry;
  viewport: ViewViewport;
  words?: (text: string) => string;
}) {
  const id = useId().replaceAll(':', ''),
    left = geometry.rowHeaderWidth,
    top = geometry.columnHeaderHeight,
    width = Math.max(0, viewport.width - left),
    height = Math.max(0, viewport.height - top);
  return (
    <svg className="sheet-page-boundary-overlay" aria-hidden="true" width={viewport.width} height={viewport.height}>
      <defs>
        <clipPath id={id}>
          <rect x={left} y={top} width={width} height={height} />
        </clipPath>
      </defs>
      <g clipPath={`url(#${id})`}>
        {model.pages.map((page) => {
          const box = projectPageBoundary(page, geometry, viewport);
          if (
            box.left > viewport.width ||
            box.top > viewport.height ||
            box.left + box.width < left ||
            box.top + box.height < top
          )
            return null;
          const x = Math.max(left, box.left),
            y = Math.max(top, box.top),
            right = Math.min(viewport.width, box.left + box.width),
            bottom = Math.min(viewport.height, box.top + box.height);
          return (
            <g key={page.number} data-page-boundary={page.number} data-page-range={page.range}>
              <rect x={box.left} y={box.top} width={box.width} height={box.height} />
              <text x={(x + right) / 2} y={(y + bottom) / 2} textAnchor="middle">
                {words('Page')} {page.number}
              </text>
            </g>
          );
        })}
      </g>
    </svg>
  );
}
