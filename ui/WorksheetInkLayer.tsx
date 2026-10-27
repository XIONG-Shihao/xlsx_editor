import { useEffect, useLayoutEffect, useMemo, useRef, useSyncExternalStore } from 'react';
import { WorkbookInkPreview } from './WorkbookInkPreview';
import { anchoredInkStroke } from './inkGeometry';
import { objectScreenBox } from './worksheetObjectGeometry';
import { createInkGestureController, type InkGestureFrame } from './inkGestureController';
export interface WorksheetInkLayerProps {
  frame: InkGestureFrame;
  width: number;
  height: number;
  onWheel: (dx: number, dy: number) => void;
  words?: (s: string) => string;
}
/** An explicit pointer surface above the readonly object layer, never a floating dialog. */
export function WorksheetInkLayer({
  frame,
  width,
  height,
  onWheel,
  words = (s) => s,
}: WorksheetInkLayerProps): JSX.Element {
  const controller = useMemo(() => createInkGestureController(), []),
    retirement = useRef<ReturnType<typeof setTimeout> | null>(null),
    surface = useRef<HTMLDivElement>(null);
  const state = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  useLayoutEffect(() => {
    controller.setFrame(frame);
  });
  useEffect(() => {
    if (retirement.current !== null) clearTimeout(retirement.current);
    return () => {
      controller.setFrame(null);
      retirement.current = setTimeout(() => controller.dispose(), 0);
    };
  }, [controller]);
  useEffect(() => {
    const cancel = () => controller.cancel(),
      escape = (e: KeyboardEvent) => {
        if (e.key === 'Escape' && frame.active) {
          e.preventDefault();
          e.stopPropagation();
          controller.cancel();
        }
      };
    window.addEventListener('blur', cancel);
    document.addEventListener('visibilitychange', cancel);
    document.addEventListener('keydown', escape, true);
    return () => {
      window.removeEventListener('blur', cancel);
      document.removeEventListener('visibilitychange', cancel);
      document.removeEventListener('keydown', escape, true);
    };
  }, [controller, frame.active]);
  let preview: null | JSX.Element = null;
  if (state.key === frame.key && state.points.length > 1) {
    try {
      const value = anchoredInkStroke(state.points, frame.pen, frame.axes),
        box = objectScreenBox(value.box, frame.view);
      preview = (
        <div
          data-ink-draft
          style={{
            position: 'absolute',
            left: box.left,
            top: box.top,
            width: box.width,
            height: box.height,
            pointerEvents: 'none',
          }}
        >
          <WorkbookInkPreview stroke={value.stroke} label={words('Draft stroke')} />
        </div>
      );
    } catch {
      /* A degenerate/edge draft is refused on release. */
    }
  }
  return (
    <div
      ref={surface}
      className="sheet-ink-layer"
      hidden={!frame.active}
      style={{
        width,
        height,
        clipPath: `inset(${frame.view.columnHeaderHeight}px 0 0 ${frame.view.rowHeaderWidth}px)`,
      }}
      role="group"
      aria-label={words('Draw surface')}
      tabIndex={0}
      onPointerDown={(e) => {
        if (e.button !== 0 || !e.isPrimary) return;
        e.preventDefault();
        e.stopPropagation();
        const b = e.currentTarget.getBoundingClientRect();
        if (controller.start(e.pointerId, e.clientX - b.left, e.clientY - b.top)) {
          e.currentTarget.focus({ preventScroll: true });
          e.currentTarget.setPointerCapture?.(e.pointerId);
        }
      }}
      onPointerMove={(e) => {
        const b = e.currentTarget.getBoundingClientRect();
        controller.move(e.pointerId, e.clientX - b.left, e.clientY - b.top);
      }}
      onPointerUp={(e) => {
        const b = e.currentTarget.getBoundingClientRect();
        void controller.finish(e.pointerId, e.clientX - b.left, e.clientY - b.top);
      }}
      onPointerCancel={() => controller.cancel()}
      onLostPointerCapture={() => controller.cancel()}
      onBlur={() => controller.cancel()}
      onWheel={(e) => {
        controller.cancel();
        onWheel(e.deltaX, e.deltaY);
      }}
    >
      {preview}
      {(state.pending || state.error || frame.reason) && (
        <p className="sheet-object-message" role="status">
          {words(state.pending ? 'Saving stroke…' : (state.error ?? frame.reason ?? ''))}
        </p>
      )}
    </div>
  );
}
