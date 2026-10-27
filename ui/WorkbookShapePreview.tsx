import type { WorkbookShapeModel } from '../core/contracts';
import { shapeSample } from '../core/shape-model.mjs';

/** Render the editable model directly, using logical pixels and fixed point text.
 * Worksheet zoom scales the SVG once; resizing changes the wrapping box. */
export function WorkbookShapePreview({
  shape,
  width = 240,
  height = 120,
  label,
}: {
  shape: WorkbookShapeModel;
  width?: number;
  height?: number;
  label?: string;
}): JSX.Element {
  const sample = shapeSample(shape.sample),
    w = Math.max(1, width),
    h = Math.max(1, height),
    stroke = sample.line ? '#' + sample.line : 'none',
    fill = sample.fill ? '#' + sample.fill : 'none';
  if (sample.family === 'icon')
    return (
      <svg
        role="img"
        aria-label={label || sample.label}
        data-shape-sample={shape.sample}
        viewBox={`0 0 ${w} ${h}`}
        style={{ width: '100%', height: '100%', display: 'block' }}
      >
        <title>{label || sample.label}</title>
        <polygon
          transform={`scale(${w / 100} ${h / 100})`}
          points={sample.points!.map((p) => p.join(',')).join(' ')}
          fill={'#' + (shape.color ?? sample.fill)}
        />
      </svg>
    );
  const style = { fill, stroke, strokeWidth: 12700 / 9525 },
    inset = sample.geometry === 'ellipse' ? 0.15 : 0,
    left = sample.geometry === 'rightArrow' ? 9.6 : w * inset + 9.6,
    right = sample.geometry === 'rightArrow' ? w * 0.35 + 9.6 : w * inset + 9.6,
    top = sample.geometry === 'rightArrow' ? h * 0.25 + 4.8 : 4.8 + h * inset,
    textHeight = sample.geometry === 'rightArrow' ? h * 0.5 - 9.6 : h - 9.6 - h * inset * 2;
  return (
    <svg
      role="img"
      aria-label={label || shape.text || sample.label}
      data-shape-sample={shape.sample}
      viewBox={`0 0 ${w} ${h}`}
      style={{ width: '100%', height: '100%', display: 'block', overflow: 'hidden' }}
    >
      <title>{label || shape.text || sample.label}</title>
      {sample.geometry === 'ellipse' ? (
        <ellipse cx={w / 2} cy={h / 2} rx={Math.max(0, w / 2 - 0.7)} ry={Math.max(0, h / 2 - 0.7)} {...style} />
      ) : sample.geometry === 'rightArrow' ? (
        <path d={`M0 ${h * 0.25}H${w * 0.65}V0L${w} ${h / 2}L${w * 0.65} ${h}V${h * 0.75}H0Z`} {...style} />
      ) : (
        <rect x={0.7} y={0.7} width={Math.max(0, w - 1.4)} height={Math.max(0, h - 1.4)} {...style} />
      )}
      <foreignObject x={left} y={top} width={Math.max(0, w - left - right)} height={Math.max(0, textHeight)}>
        <div style={{ height: '100%', display: 'flex', alignItems: 'center', overflow: 'hidden' }}>
          <div
            style={{
              width: '100%',
              whiteSpace: 'pre-wrap',
              overflowWrap: 'anywhere',
              fontFamily: 'Arial, sans-serif',
              fontSize: sample.fontSize + 'pt',
              fontWeight: sample.family === 'wordart' ? 700 : 400,
              fontStyle: 'normal',
              lineHeight: 1.2,
              color: sample.family === 'wordart' && sample.outline ? 'transparent' : '#' + sample.color,
              WebkitTextStroke: sample.outline ? '1.333333px #' + sample.outline : undefined,
              textShadow: sample.shadow ? '4px 4px 2px rgba(0,0,0,0.35)' : undefined,
              textAlign: sample.align === 'ctr' ? 'center' : 'left',
            }}
          >
            {shape.text}
          </div>
        </div>
      </foreignObject>
    </svg>
  );
}
