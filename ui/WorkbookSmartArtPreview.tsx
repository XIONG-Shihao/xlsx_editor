import { useId } from 'react';
import type { WorkbookSmartArtModel } from '../core/contracts';
import { smartartGeometry } from '../core/smartart-model.mjs';
export function WorkbookSmartArtPreview({
  model,
  width = 600,
  height = 360,
  label,
}: {
  model: WorkbookSmartArtModel;
  width?: number;
  height?: number;
  label?: string;
}): JSX.Element {
  const marker = useId().replaceAll(':', ''),
    g = smartartGeometry(model),
    x = (v: number) => (v * width) / g.width,
    y = (v: number) => (v * height) / g.height;
  return (
    <svg
      role="img"
      aria-label={label || 'SmartArt ' + model.layout}
      data-smartart-layout={model.layout}
      viewBox={`0 0 ${width} ${height}`}
      style={{ width: '100%', height: '100%', display: 'block' }}
    >
      <title>{label || 'SmartArt ' + model.layout}</title>
      <defs>
        <marker id={marker} markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto">
          <path d="M0 0L8 4L0 8Z" fill="#4472C4" />
        </marker>
      </defs>
      {g.edges.map((e, i) => (
        <line
          key={i}
          data-smartart-connection={e.from + '-' + e.to}
          x1={x(e.x1)}
          y1={y(e.y1)}
          x2={x(e.x2)}
          y2={y(e.y2)}
          stroke="#4472C4"
          strokeWidth={2}
          markerEnd={'url(#' + marker + ')'}
        />
      ))}
      {g.boxes.map((b) => (
        <g key={b.id} data-smartart-node={b.id}>
          <rect
            x={x(b.x)}
            y={y(b.y)}
            width={x(b.w)}
            height={y(b.h)}
            rx={8}
            fill="#4472C4"
            stroke="#2F5597"
            strokeWidth={12700 / 9525}
          />
          <foreignObject x={x(b.x) + 4} y={y(b.y) + 4} width={Math.max(0, x(b.w) - 8)} height={Math.max(0, y(b.h) - 8)}>
            <div
              style={{
                height: '100%',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                textAlign: 'center',
                fontFamily: 'Arial, sans-serif',
                fontSize: '14pt',
                color: 'white',
                overflow: 'hidden',
                overflowWrap: 'anywhere',
              }}
            >
              {b.text}
            </div>
          </foreignObject>
        </g>
      ))}
    </svg>
  );
}
