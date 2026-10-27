import { inkPen, inkStrokePath, normalizeInkStroke, type InkStrokeModel } from '../core/ink-stroke.mjs';
/** Direct open-vector rendering, with one view scale applied by its worksheet container. */
export function WorkbookInkPreview({ stroke, label }: { stroke: InkStrokeModel; label?: string }): JSX.Element {
  const model = normalizeInkStroke(stroke),
    pen = inkPen(model.pen),
    title = label || pen.label + ' stroke';
  return (
    <svg
      role="img"
      aria-label={title}
      data-ink-pen={pen.id}
      viewBox={`0 0 ${model.widthEmu} ${model.heightEmu}`}
      preserveAspectRatio="none"
      style={{ width: '100%', height: '100%', display: 'block', overflow: 'hidden' }}
    >
      <title>{title}</title>
      <path
        d={inkStrokePath(model)}
        fill="none"
        stroke={'#' + pen.color}
        strokeWidth={pen.widthEmu}
        strokeOpacity={pen.alpha / 100000}
        strokeLinecap={pen.cap === 'round' ? 'round' : 'butt'}
        strokeLinejoin="round"
      />
    </svg>
  );
}
