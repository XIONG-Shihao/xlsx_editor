import { useEffect, useMemo, useState } from 'react';
import {
  normalGeometryFont,
  referenceGeometryProfile,
  validateGeometryProfile,
  type GeometryFont,
  type GeometryProfile,
  type WorksheetGeometrySource,
} from '../core/worksheet-geometry.mjs';

/** Quote the accepted font family as a single CSS family, never CSS source. */
export function geometryFontFamily(font: GeometryFont): string {
  return JSON.stringify(font.fontFamily) + ', sans-serif';
}
export function geometryCanvasFont(font: GeometryFont, fallback = 'sans-serif'): string {
  return `${font.italic ? 'italic' : 'normal'} ${font.bold ? 'bold' : 'normal'} ${(font.fontSize * 96) / 72}px ${JSON.stringify(font.fontFamily)}, ${fallback}`;
}
export function measureGeometryProfile(
  font: GeometryFont,
  context?: Pick<CanvasRenderingContext2D, 'font' | 'measureText'> | null
): GeometryProfile {
  if (!context) {
    try {
      if (typeof document !== 'undefined' && typeof CanvasRenderingContext2D !== 'undefined')
        context = document.createElement('canvas').getContext('2d');
    } catch {
      /* Browser/SSR without a canvas uses a named fallback. */
    }
  }
  if (!context) return referenceGeometryProfile(font);
  try {
    // A missing requested family falls back differently for serif/monospace.
    // This detects substitution but matching measurements alone do not confirm
    // installed-family availability, so local measurements remain unverified.
    context.font = geometryCanvasFont(font, 'monospace');
    const first = context.measureText('0123456789Wiil').width;
    context.font = geometryCanvasFont(font, 'serif');
    const second = context.measureText('0123456789Wiil').width;
    const substituted = Math.abs(first - second) > 0.0001;
    context.font = substituted
      ? `${font.italic ? 'italic' : 'normal'} ${font.bold ? 'bold' : 'normal'} ${(font.fontSize * 96) / 72}px sans-serif`
      : geometryCanvasFont(font);
    const maximum = Math.max(...Array.from('0123456789', (digit) => context!.measureText(digit).width));
    if (!Number.isFinite(maximum) || maximum <= 0 || maximum > 4096) throw Error('Invalid font measurement');
    return validateGeometryProfile(
      {
        schema: 1,
        font,
        maximumDigitWidth: Math.max(1, Math.round(maximum)),
        source: substituted ? 'substituted-font' : 'local-font',
        availability: substituted ? 'substituted' : 'unverified',
        ...(substituted ? { renderedFontFamily: 'sans-serif' } : {}),
      },
      font
    );
  } catch {
    return referenceGeometryProfile(font);
  }
}
/** Metric updates are view-only. A late font load cannot adopt another font's
 * measurements, and the new font is masked synchronously before an effect. */
export function useWorksheetGeometryProfile(sheet?: WorksheetGeometrySource): GeometryProfile {
  const font = normalGeometryFont(sheet),
    key = JSON.stringify(font);
  const initial = useMemo(() => measureGeometryProfile(font), [key]);
  const [loaded, setLoaded] = useState({ key, profile: initial });
  useEffect(() => {
    let live = true;
    const refresh = () => {
      if (live) {
        const profile = measureGeometryProfile(font);
        setLoaded((previous) =>
          previous.key === key && JSON.stringify(previous.profile) === JSON.stringify(profile)
            ? previous
            : { key, profile }
        );
      }
    };
    const fonts = typeof document === 'undefined' ? undefined : document.fonts;
    void fonts?.ready.then(refresh).catch(() => {});
    fonts?.addEventListener('loadingdone', refresh);
    return () => {
      live = false;
      fonts?.removeEventListener('loadingdone', refresh);
    };
  }, [key]);
  return loaded.key === key ? loaded.profile : initial;
}
