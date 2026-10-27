import { useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import type { WorkbookObjectPreview, WorkbookObjectCommand } from '../core/contracts';
import type { WorkbookObjectsProps } from './WorkbookObjects';
import { WorkbookObjectPreviewView } from './WorkbookObjects';
import { createWorksheetObjectController, type WorksheetObjectFrame } from './worksheetObjectController';
import { inkWorksheetPoint } from './inkGeometry';
import { pickOwnedInk } from './inkSelection';
import {
  objectGestureBox,
  objectPlacementCommand,
  objectScreenBox,
  type ObjectAxes,
  type ObjectBox,
} from './worksheetObjectGeometry';
export interface WorksheetObjectLayerProps {
  inkMode?: 'select' | 'erase' | null;
  keyValue: string;
  family: string;
  gestureKey: string;
  sheet: string;
  active: boolean;
  reason?: string;
  controls: Pick<WorkbookObjectsProps, 'objects' | 'object' | 'onBatch'>;
  axes: ObjectAxes;
  view: {
    scale: number;
    rowHeaderWidth: number;
    columnHeaderHeight: number;
    scrollLeft: number;
    scrollTop: number;
    width: number;
    height: number;
  };
  onExtent: (family: string, extent: { rows: number; columns: number }) => void;
  words?: (text: string) => string;
}
interface Gesture {
  frame: WorksheetObjectFrame;
  gestureKey: string;
  pointer: number;
  startX: number;
  startY: number;
  scale: number;
  kind: 'move' | 'resize';
  object: WorkbookObjectPreview;
  original: ObjectBox;
  draft: ObjectBox;
}
export function WorksheetObjectLayer(props: WorksheetObjectLayerProps): JSX.Element {
  const words = props.words ?? ((text: string) => text),
    controller = useMemo(() => createWorksheetObjectController(), []),
    disposal = useRef<ReturnType<typeof setTimeout> | null>(null);
  const frame: WorksheetObjectFrame = {
      key: props.keyValue,
      family: props.family,
      sheet: props.sheet,
      active: props.active,
      reason: props.reason,
      controls: { ...props.controls },
    },
    live = useRef(props);
  live.current = props;
  const currentFrame = useRef(frame);
  currentFrame.current = frame;
  const snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const fresh = snapshot.key === frame.key && frame.active,
    objects = fresh ? snapshot.objects : [],
    pending = snapshot.pending;
  const [selection, setSelection] = useState<{ family: string; id: string } | null>(null),
    [draft, setDraft] = useState<{ id: string; key: string; box: ObjectBox } | null>(null),
    [notice, setNotice] = useState<{ key: string; text: string } | null>(null);
  const inkGesture = useRef<(Omit<Gesture, 'kind'> & { kind: 'move' | 'erase' }) | null>(null);
  const gesture = useRef<Gesture | null>(null),
    selected = selection?.family === props.family ? selection.id : null;
  useLayoutEffect(() => {
    controller.setFrame(frame);
  });
  useEffect(() => {
    if (disposal.current !== null) clearTimeout(disposal.current);
    return () => {
      controller.setFrame(null);
      disposal.current = setTimeout(() => controller.dispose(), 0);
    };
  }, [controller]);
  useEffect(() => {
    if (
      gesture.current &&
      (gesture.current.frame.key !== frame.key ||
        gesture.current.gestureKey !== props.gestureKey ||
        props.reason ||
        !props.active)
    ) {
      gesture.current = null;
      setDraft(null);
    }
  }, [frame.key, props.gestureKey, props.reason, props.active]);
  useEffect(() => {
    if (!fresh) return;
    try {
      let rows = 30,
        columns = 12;
      for (const item of objects)
        if (item.anchor) {
          const e = props.axes.extent(props.axes.box(item.anchor));
          rows = Math.max(rows, e.rows);
          columns = Math.max(columns, e.columns);
        }
      props.onExtent(props.family, { rows, columns });
    } catch (error) {
      setNotice({ key: frame.key, text: String(error instanceof Error ? error.message : error) });
    }
  }, [fresh, objects, props.axes, props.family, props.onExtent, frame.key]);
  const selectionLive = useRef<string | null>(selected);
  selectionLive.current = selected;
  const cancel = () => {
    gesture.current = null;
    inkGesture.current = null;
    setDraft(null);
  };
  useEffect(() => {
    if (inkGesture.current && selected !== inkGesture.current.object.id) cancel();
  }, [selected]);
  useEffect(() => {
    if (
      inkGesture.current &&
      (inkGesture.current.frame.key !== frame.key ||
        inkGesture.current.gestureKey !== props.gestureKey ||
        props.reason ||
        !props.active ||
        !props.inkMode)
    )
      cancel();
  }, [frame.key, props.gestureKey, props.reason, props.active, props.inkMode]);
  useEffect(() => {
    const clear = () => cancel(),
      escape = (e: KeyboardEvent) => {
        if (e.key === 'Escape' && props.inkMode && props.active) {
          e.preventDefault();
          e.stopPropagation();
          cancel();
          setSelection(null);
        }
      };
    window.addEventListener('blur', clear);
    document.addEventListener('visibilitychange', clear);
    document.addEventListener('keydown', escape, true);
    return () => {
      window.removeEventListener('blur', clear);
      document.removeEventListener('visibilitychange', clear);
      document.removeEventListener('keydown', escape, true);
    };
  }, [props.inkMode, props.active]);
  useEffect(() => {
    const clear = (event: PointerEvent) => {
      if (event.target instanceof Element && event.target.closest('.sheet-grid')) {
        gesture.current = null;
        setDraft(null);
        setSelection(null);
      }
    };
    document.addEventListener('pointerdown', clear);
    return () => document.removeEventListener('pointerdown', clear);
  }, []);
  async function submit(captured: WorksheetObjectFrame, command: WorkbookObjectCommand) {
    if (currentFrame.current.key !== captured.key) return;
    try {
      const result = await controller.accept(captured, command);
      if (result.current) {
        setNotice(null);
        controller.refresh();
      }
    } catch (error) {
      if (currentFrame.current.key === captured.key)
        setNotice({ key: captured.key, text: String(error instanceof Error ? error.message : error) });
    }
  }
  function start(event: React.PointerEvent<HTMLElement>, object: WorkbookObjectPreview, kind: 'move' | 'resize') {
    if (event.button !== 0 || currentFrame.current.key !== frame.key) return;
    const input = document.activeElement;
    if (
      input instanceof HTMLElement &&
      (input.getAttribute('data-editing') === 'true' || input.getAttribute('aria-label') === 'Formula or value')
    ) {
      event.preventDefault();
      setNotice({ key: frame.key, text: 'Finish or cancel the cell edit before selecting a drawing.' });
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    setSelection({ family: props.family, id: object.id });
    event.currentTarget.focus();
    if (object.kind === 'ink' && !props.inkMode) return;
    if (props.reason || pending || !object.canEdit || object.protectedReason || !object.anchor) return;
    try {
      const original = props.axes.box(object.anchor);
      gesture.current = {
        frame: { ...frame, controls: { ...frame.controls } },
        gestureKey: props.gestureKey,
        pointer: event.pointerId,
        startX: event.clientX,
        startY: event.clientY,
        scale: props.view.scale,
        kind,
        object: structuredClone(object),
        original,
        draft: original,
      };
      setNotice(null);
      event.currentTarget.setPointerCapture?.(event.pointerId);
    } catch (error) {
      setNotice({ key: frame.key, text: String(error instanceof Error ? error.message : error) });
    }
  }
  function move(event: React.PointerEvent) {
    const g = gesture.current;
    if (!g || g.pointer !== event.pointerId) return;
    event.preventDefault();
    event.stopPropagation();
    try {
      g.draft = objectGestureBox(
        g.original,
        g.kind,
        (event.clientX - g.startX) / g.scale,
        (event.clientY - g.startY) / g.scale,
        g.object
      );
      live.current.axes.anchor(g.draft);
      setDraft({ id: g.object.id, key: g.frame.key, box: g.draft });
    } catch (error) {
      cancel();
      setNotice({ key: g.frame.key, text: String(error instanceof Error ? error.message : error) });
    }
  }
  function finish(event: React.PointerEvent) {
    const g = gesture.current;
    if (!g || g.pointer !== event.pointerId) return;
    move(event);
    if (gesture.current !== g) return;
    cancel();
    event.preventDefault();
    event.stopPropagation();
    if (
      currentFrame.current.key !== g.frame.key ||
      live.current.gestureKey !== g.gestureKey ||
      JSON.stringify(g.draft) === JSON.stringify(g.original)
    )
      return;
    try {
      void submit(g.frame, objectPlacementCommand(g.frame.sheet, g.object, live.current.axes.anchor(g.draft)));
    } catch (error) {
      setNotice({ key: g.frame.key, text: String(error instanceof Error ? error.message : error) });
    }
  }
  function keys(event: React.KeyboardEvent, object: WorkbookObjectPreview) {
    if ((event.metaKey || event.ctrlKey) && ['c', 'x', 'v', 'a'].includes(event.key.toLowerCase())) {
      event.preventDefault();
      event.stopPropagation();
      setNotice({
        key: frame.key,
        text: 'Object clipboard and multiple selection are not supported. Use the drawing controls.',
      });
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      cancel();
      setSelection(null);
      return;
    }
    if (object.kind === 'ink' && !props.inkMode) return;
    if (!['Delete', 'Backspace', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return;
    event.preventDefault();
    event.stopPropagation();
    if (event.repeat || props.reason || pending || gesture.current || inkGesture.current) return;
    if (event.key === 'Delete' || event.key === 'Backspace') {
      void submit(
        { ...frame, controls: { ...frame.controls } },
        { operation: 'deleteObject', sheet: props.sheet, objectId: object.id, expectedObjectRevision: object.revision }
      );
      return;
    }
    if (!object.anchor) return;
    try {
      const step = event.shiftKey ? 10 : 1,
        dx = event.key === 'ArrowRight' ? step : event.key === 'ArrowLeft' ? -step : 0,
        dy = event.key === 'ArrowDown' ? step : event.key === 'ArrowUp' ? -step : 0,
        box = objectGestureBox(props.axes.box(object.anchor), 'move', dx, dy, object);
      void submit(
        { ...frame, controls: { ...frame.controls } },
        objectPlacementCommand(props.sheet, object, props.axes.anchor(box))
      );
    } catch (error) {
      setNotice({ key: frame.key, text: String(error instanceof Error ? error.message : error) });
    }
  }
  function inkStart(event: React.PointerEvent<HTMLDivElement>) {
    if (
      event.button !== 0 ||
      !event.isPrimary ||
      props.reason ||
      pending ||
      !fresh ||
      inkGesture.current ||
      currentFrame.current.key !== frame.key
    )
      return;
    const b = event.currentTarget.getBoundingClientRect(),
      point = inkWorksheetPoint(props.view, event.clientX - b.left, event.clientY - b.top);
    if (!point) return;
    const object = pickOwnedInk(objects, point, props.axes, props.view.scale);
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.focus({ preventScroll: true });
    setSelection(object ? { family: props.family, id: object.id } : null);
    selectionLive.current = object?.id ?? null;
    if (!object?.anchor) return;
    const original = props.axes.box(object.anchor);
    inkGesture.current = {
      frame: { ...frame, controls: { ...frame.controls } },
      gestureKey: props.gestureKey,
      pointer: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      scale: props.view.scale,
      kind: props.inkMode === 'erase' ? 'erase' : 'move',
      object: structuredClone(object),
      original,
      draft: original,
    };
    setNotice(null);
    event.currentTarget.setPointerCapture?.(event.pointerId);
  }
  function inkMove(event: React.PointerEvent<HTMLDivElement>) {
    const g = inkGesture.current;
    if (!g || g.pointer !== event.pointerId) return;
    if (
      currentFrame.current.key !== g.frame.key ||
      live.current.gestureKey !== g.gestureKey ||
      selectionLive.current !== g.object.id ||
      live.current.reason
    ) {
      cancel();
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    if (g.kind === 'erase') return;
    try {
      g.draft = objectGestureBox(
        g.original,
        'move',
        (event.clientX - g.startX) / g.scale,
        (event.clientY - g.startY) / g.scale,
        g.object
      );
      props.axes.anchor(g.draft);
      setDraft({ id: g.object.id, key: g.frame.key, box: g.draft });
    } catch (e) {
      cancel();
      setNotice({ key: g.frame.key, text: String(e instanceof Error ? e.message : e) });
    }
  }
  function inkFinish(event: React.PointerEvent<HTMLDivElement>) {
    const g = inkGesture.current;
    if (!g || g.pointer !== event.pointerId) return;
    inkMove(event);
    if (inkGesture.current !== g) return;
    cancel();
    if (
      currentFrame.current.key !== g.frame.key ||
      live.current.gestureKey !== g.gestureKey ||
      selectionLive.current !== g.object.id ||
      live.current.reason
    )
      return;
    if (g.kind === 'erase') {
      void submit(g.frame, {
        operation: 'deleteInkStroke',
        sheet: g.frame.sheet,
        objectId: g.object.id,
        expectedObjectRevision: g.object.revision,
      });
      return;
    }
    if (JSON.stringify(g.draft) !== JSON.stringify(g.original))
      void submit(g.frame, objectPlacementCommand(g.frame.sheet, g.object, props.axes.anchor(g.draft)));
  }
  return (
    <div
      className="sheet-object-layer"
      aria-label={words('Worksheet objects')}
      hidden={!props.active}
      style={{ clipPath: `inset(${props.view.columnHeaderHeight}px 0 0 ${props.view.rowHeaderWidth}px)` }}
    >
      {objects.map((object) => {
        if (!object.anchor) return null;
        let box: ObjectBox;
        try {
          box = objectScreenBox(
            draft?.key === frame.key && draft.id === object.id ? draft.box : props.axes.box(object.anchor),
            props.view
          );
        } catch {
          return null;
        }
        if (
          box.left + box.width <= props.view.rowHeaderWidth ||
          box.top + box.height <= props.view.columnHeaderHeight ||
          box.left >= props.view.width ||
          box.top >= props.view.height
        )
          return null;
        const chosen = selected === object.id,
          reason =
            props.reason ??
            object.protectedReason ??
            (object.kind === 'ink' && !props.inkMode
              ? 'Choose Draw → Select Strokes or Eraser to edit this stroke.'
              : undefined) ??
            (!object.canEdit ? 'This object is preserved and protected.' : undefined);
        return (
          <div
            key={object.id}
            data-worksheet-object={object.id}
            data-object-kind={object.kind}
            data-object-selected={chosen}
            className="sheet-worksheet-object"
            style={{ left: box.left, top: box.top, width: box.width, height: box.height }}
            tabIndex={0}
            role="group"
            aria-label={words('Worksheet object') + ': ' + object.name}
            aria-roledescription={words('drawing object')}
            title={reason ? words(reason) : object.name}
            onFocus={() => setSelection({ family: props.family, id: object.id })}
            onPointerDown={(event) => start(event, object, 'move')}
            onPointerMove={move}
            onPointerUp={finish}
            onPointerCancel={cancel}
            onLostPointerCapture={cancel}
            onKeyDown={(event) => keys(event, object)}
          >
            <div className="sheet-object-content">
              <WorkbookObjectPreviewView object={object} embedded words={words} />
            </div>
            {chosen && (
              <>
                <button
                  type="button"
                  className="sheet-object-delete"
                  aria-label={words('Delete object') + ': ' + object.name}
                  disabled={!!reason || !object.canDelete || pending}
                  onPointerDown={(event) => event.stopPropagation()}
                  onClick={(event) => {
                    event.stopPropagation();
                    void submit(
                      { ...frame, controls: { ...frame.controls } },
                      {
                        operation: 'deleteObject',
                        sheet: props.sheet,
                        objectId: object.id,
                        expectedObjectRevision: object.revision,
                      }
                    );
                  }}
                >
                  ×
                </button>
                <button
                  type="button"
                  className="sheet-object-resize"
                  aria-label={words('Resize object') + ': ' + object.name}
                  disabled={!!reason || !object.canEdit || object.kind === 'ink' || pending}
                  onPointerDown={(event) => start(event, object, 'resize')}
                  onPointerMove={move}
                  onPointerUp={finish}
                  onPointerCancel={cancel}
                  onLostPointerCapture={cancel}
                  onKeyDown={(event) => keys(event, object)}
                />
              </>
            )}
          </div>
        );
      })}
      {props.inkMode && props.active && (
        <div
          className="sheet-ink-edit-surface"
          role="group"
          aria-label={words(props.inkMode === 'erase' ? 'Whole-stroke eraser' : 'Select strokes')}
          tabIndex={0}
          style={{
            pointerEvents: 'auto',
            position: 'absolute',
            inset: 0,
            touchAction: 'none',
            cursor: props.inkMode === 'erase' ? 'crosshair' : 'default',
          }}
          onPointerDown={inkStart}
          onPointerMove={inkMove}
          onPointerUp={inkFinish}
          onPointerCancel={cancel}
          onLostPointerCapture={cancel}
          onBlur={cancel}
          onKeyDown={(event) => {
            const object = objects.find((o) => o.id === selected);
            if (object?.kind === 'ink') keys(event, object);
          }}
        >
          {props.inkMode === 'select' &&
            selected &&
            objects
              .filter((o) => o.id === selected && o.kind === 'ink')
              .map((o) => {
                if (!o.anchor) return null;
                const b = objectScreenBox(props.axes.box(o.anchor), props.view);
                return (
                  <button
                    key={o.id}
                    type="button"
                    className="sheet-object-delete"
                    aria-label={words('Delete stroke')}
                    disabled={!!props.reason || pending || !o.canDelete}
                    style={{ left: b.left + b.width, top: b.top, right: 'auto' }}
                    onPointerDown={(e) => e.stopPropagation()}
                    onClick={() =>
                      void submit(
                        { ...frame, controls: { ...frame.controls } },
                        {
                          operation: 'deleteInkStroke',
                          sheet: props.sheet,
                          objectId: o.id,
                          expectedObjectRevision: o.revision,
                        }
                      )
                    }
                  >
                    ×
                  </button>
                );
              })}
        </div>
      )}
      {((fresh && snapshot.error) || notice?.key === frame.key || (fresh && pending)) && (
        <p className="sheet-object-message" role="status">
          {words(notice?.key === frame.key ? notice.text : snapshot.error || 'Applying object change…')}
        </p>
      )}
    </div>
  );
}
