import type { WorkbookModel3DAsset } from '../core/contracts';
import { MODEL3D_SAMPLES, projectModel3D } from '../core/model3d-model.mjs';
/** Software mesh projection renders actual decoded vertices; no image/cache thumbnail. */
export function WorkbookModel3DPreview({
  asset,
  width = 320,
  height = 240,
  label,
}: {
  asset: WorkbookModel3DAsset;
  width?: number;
  height?: number;
  label?: string;
}): JSX.Element {
  const frame = projectModel3D(asset, width, height),
    name = label || MODEL3D_SAMPLES.find((s) => s.id === asset.model.sample)!.label;
  return (
    <svg
      role="img"
      aria-label={name}
      data-model3d-sample={asset.model.sample}
      viewBox={`0 0 ${width} ${height}`}
      style={{ width: '100%', height: '100%', display: 'block' }}
    >
      <title>{name}</title>
      {frame.faces.map((face) => (
        <polygon
          key={face.index}
          data-model3d-face={face.index}
          points={face.points.map((p) => p.join(',')).join(' ')}
          fill={face.color}
          stroke={face.color}
          strokeWidth={0.5}
          strokeLinejoin="round"
        />
      ))}
    </svg>
  );
}
