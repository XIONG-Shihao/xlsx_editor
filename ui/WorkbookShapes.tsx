import { useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import type { WorkbookObjectCommand, WorkbookObjectPreview, WorkbookShapeSample } from '../core/contracts';
import { ILLUSTRATION_SAMPLES, SHAPE_LIMITS, normalizeShape, shapeSample } from '../core/shape-model.mjs';
import {
  createWorksheetObjectController,
  type WorksheetObjectFrame,
  type WorksheetObjectController,
} from './worksheetObjectController';
import { WorkbookShapePreview } from './WorkbookShapePreview';
export type IllustrationMode = 'text' | 'shape' | 'icon' | 'wordart' | 'smartart' | 'model3d';
export function WorkbookShapes({
  frame,
  mode,
  selection,
  controller: sharedController,
  words = (s: string) => s,
}: {
  frame: WorksheetObjectFrame;
  mode: Exclude<IllustrationMode, 'smartart' | 'model3d'>;
  selection: string;
  controller?: WorksheetObjectController;
  words?: (s: string) => string;
}): JSX.Element {
  const controller = useMemo(() => sharedController ?? createWorksheetObjectController(), [sharedController]),
    snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot),
    live = useRef(frame);
  live.current = frame;
  useLayoutEffect(() => controller.setFrame(frame), [controller, frame.key, frame.active, frame.reason]);
  const disposal = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(() => {
    if (sharedController) return;
    if (disposal.current) clearTimeout(disposal.current);
    return () => {
      disposal.current = setTimeout(() => controller.dispose(), 0);
    };
  }, [controller, sharedController]);
  const first =
      mode === 'text'
        ? 'text-plain'
        : mode === 'icon'
          ? 'icon-check'
          : mode === 'wordart'
            ? 'wordart-blue'
            : 'shape-rectangle',
    defaultText = mode === 'icon' ? '' : mode === 'wordart' ? 'WordArt' : shapeSample(first).label,
    defaultWidth = mode === 'icon' ? 96 : mode === 'wordart' ? 320 : 240,
    defaultHeight = mode === 'icon' ? 96 : 120,
    defaultCell = selection.split(':')[0],
    fresh = snapshot.key === frame.key,
    objects = fresh ? snapshot.objects : [],
    pending = snapshot.pending,
    disabled = !frame.active || !!frame.reason || pending || !fresh || snapshot.loading || !!snapshot.error;
  const [draft, setDraft] = useState({
      key: frame.key,
      id: null as string | null,
      sample: first as WorkbookShapeSample,
      text: defaultText,
      cell: defaultCell,
      width: defaultWidth,
      height: defaultHeight,
      alt: '',
      color: '217346',
      dx: 0,
      dy: 0,
    }),
    [notice, setNotice] = useState({ key: '', text: '' });
  const current =
    draft.key === frame.key
      ? draft
      : {
          key: frame.key,
          id: null,
          sample: first as WorkbookShapeSample,
          text: defaultText,
          cell: defaultCell,
          width: defaultWidth,
          height: defaultHeight,
          alt: '',
          color: '217346',
          dx: 0,
          dy: 0,
        };
  const selected = current.id ? objects.find((o) => o.id === current.id) : null;
  const update = (patch: Partial<typeof draft>) => {
    if (live.current.key === frame.key && frame.active) setDraft({ ...current, ...patch });
  };
  const choose = (object: WorkbookObjectPreview) => {
    if (!object.shape || !object.anchor) return;
    update({
      id: object.id,
      sample: object.shape.sample,
      text: object.shape.text,
      cell: object.anchor.cell,
      width: object.anchor.widthEmu / 9525,
      height: object.anchor.heightEmu / 9525,
      alt: object.altText ?? '',
      color: object.shape.color ?? '217346',
      dx: object.anchor.dxEmu,
      dy: object.anchor.dyEmu,
    });
  };
  const apply = async (command: WorkbookObjectCommand) => {
    if (live.current.key !== frame.key || !frame.active || disabled) return;
    try {
      setNotice({ key: frame.key, text: '' });
      const result = await controller.accept({ ...frame, controls: { ...frame.controls } }, command);
      if (result.current) {
        setDraft({ ...current, id: null });
        controller.refresh();
      }
    } catch (error) {
      if (live.current.key === frame.key)
        setNotice({ key: frame.key, text: String(error instanceof Error ? error.message : error) });
    }
  };
  const reset = () =>
    update({
      id: null,
      sample: first,
      text: defaultText,
      cell: defaultCell,
      width: defaultWidth,
      height: defaultHeight,
      alt: '',
      color: '217346',
      dx: 0,
      dy: 0,
    });
  return (
    <section
      aria-label={words(mode === 'icon' ? 'Icons' : mode === 'wordart' ? 'WordArt' : 'Text boxes and shapes')}
      style={{ padding: 12, display: 'grid', gap: 12 }}
    >
      <p>
        {words(
          mode === 'icon'
            ? 'Choose a vector sample and color, then insert. Select a saved icon below to edit it. Move and resize on the worksheet.'
            : 'Choose a sample, enter text, then insert. Select a saved object below to edit it. Move and resize on the worksheet.'
        )}
      </p>
      <div
        role="group"
        aria-label={words('Sample gallery')}
        style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(110px,1fr))', gap: 8 }}
      >
        {ILLUSTRATION_SAMPLES.filter((s) => s.family === mode).map((s) => (
          <button
            type="button"
            key={s.id}
            aria-label={words(s.label)}
            aria-pressed={current.sample === s.id}
            disabled={disabled}
            onClick={() =>
              update({
                sample: s.id,
                text: mode === 'icon' ? '' : mode === 'wordart' ? current.text : current.id ? current.text : s.label,
              })
            }
            style={{ padding: 8 }}
          >
            <span style={{ display: 'block', height: 60, pointerEvents: 'none' }}>
              <WorkbookShapePreview
                shape={{
                  sample: s.id,
                  text: mode === 'icon' ? '' : mode === 'wordart' ? 'WordArt' : s.label,
                  ...(mode === 'icon' ? { color: current.color } : {}),
                }}
                width={mode === 'icon' ? 96 : mode === 'wordart' ? 320 : 240}
                height={mode === 'icon' ? 96 : 120}
              />
            </span>
            {words(s.label)}
          </button>
        ))}
      </div>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          try {
            const shape = normalizeShape({
                sample: current.sample,
                text: current.text,
                ...(mode === 'icon' ? { color: current.color } : {}),
              }),
              anchor = {
                cell: current.cell,
                dxEmu: current.dx,
                dyEmu: current.dy,
                widthEmu: Math.round(current.width * 9525),
                heightEmu: Math.round(current.height * 9525),
              };
            if (
              !Number.isFinite(current.width) ||
              !Number.isFinite(current.height) ||
              anchor.widthEmu < 1 ||
              anchor.heightEmu < 1 ||
              anchor.widthEmu > 50000000 ||
              anchor.heightEmu > 50000000
            )
              throw Error('Choose a supported width and height.');
            void apply(
              selected
                ? {
                    operation: 'updateShape',
                    sheet: frame.sheet,
                    objectId: selected.id,
                    expectedObjectRevision: selected.revision,
                    patch: shape,
                    anchor,
                    altText: current.alt,
                  }
                : {
                    operation: 'createShape',
                    sheet: frame.sheet,
                    objectId: crypto.randomUUID(),
                    expectedObjectRevision: null,
                    shape,
                    anchor,
                    altText: current.alt,
                  }
            );
          } catch (error) {
            setNotice({ key: frame.key, text: String(error instanceof Error ? error.message : error) });
          }
        }}
      >
        <fieldset disabled={disabled || (!!selected && !selected.canEdit)} style={{ display: 'grid', gap: 10 }}>
          <legend>
            {words(
              mode === 'icon'
                ? selected
                  ? 'Edit icon'
                  : 'Insert icon'
                : mode === 'wordart'
                  ? selected
                    ? 'Edit WordArt'
                    : 'Insert WordArt'
                  : selected
                    ? 'Edit text or shape'
                    : 'Insert text or shape'
            )}
          </legend>
          {mode !== 'icon' && (
            <label>
              {words('Object text')}
              <textarea
                aria-label={words('Object text')}
                maxLength={mode === 'wordart' ? 256 : SHAPE_LIMITS.text}
                rows={3}
                value={current.text}
                onChange={(e) => update({ text: e.target.value })}
              />
            </label>
          )}
          {mode === 'icon' && (
            <label>
              {words('Icon color')}
              <input
                type="color"
                aria-label={words('Icon color')}
                value={'#' + current.color}
                onChange={(e) => update({ color: e.target.value.slice(1).toUpperCase() })}
              />
            </label>
          )}
          <label>
            {words('Anchor cell')}
            <input
              aria-label={words('Object anchor')}
              value={current.cell}
              required
              onChange={(e) => update({ cell: e.target.value, dx: 0, dy: 0 })}
            />
          </label>
          <label>
            {words('Width in pixels')}
            <input
              aria-label={words('Object width')}
              type="number"
              min="1"
              max="5249"
              step="any"
              required
              value={current.width}
              onChange={(e) => update({ width: Number(e.target.value) })}
            />
          </label>
          <label>
            {words('Height in pixels')}
            <input
              aria-label={words('Object height')}
              type="number"
              min="1"
              max="5249"
              step="any"
              required
              value={current.height}
              onChange={(e) => update({ height: Number(e.target.value) })}
            />
          </label>
          <label>
            {words('Alternative text')}
            <input
              aria-label={words('Object alternative text')}
              maxLength={SHAPE_LIMITS.alt}
              value={current.alt}
              onChange={(e) => update({ alt: e.target.value })}
            />
          </label>
          <p>
            {words(
              mode === 'icon'
                ? 'Three repository-authored vector samples with editable solid color.'
                : mode === 'wordart'
                  ? 'Arial 32 pt bold; three fill/outline/shadow presets. Up to 256 characters and four paragraphs.'
                  : 'Arial 14 pt; fixed sample colors. Up to 1,024 characters and eight paragraphs.'
            )}
          </p>
          <div style={{ height: 120, pointerEvents: 'none' }}>
            <WorkbookShapePreview
              shape={{
                sample: current.sample,
                text: current.text,
                ...(mode === 'icon' ? { color: current.color } : {}),
              }}
              width={current.width}
              height={current.height}
              label={current.alt}
            />
          </div>
          <button type="submit">{words(selected ? 'Apply object changes' : 'Insert object')}</button>
          <button type="button" onClick={reset}>
            {words('Reset form')}
          </button>
        </fieldset>
      </form>
      {frame.reason && <p role="status">{words(frame.reason)}</p>}
      {fresh && snapshot.loading && <p role="status">{words('Loading objects…')}</p>}
      {pending && <p role="status">{words('Applying object change…')}</p>}
      {((fresh && snapshot.error) || (notice.key === frame.key && notice.text)) && (
        <p role="alert">{(fresh && snapshot.error) || notice.text}</p>
      )}
      <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
        {objects
          .filter((o) => o.shape && shapeSample(o.shape.sample).family === mode)
          .map((o) => (
            <li key={o.id}>
              <button
                type="button"
                data-shape-object={o.id}
                aria-pressed={selected?.id === o.id}
                disabled={pending}
                onClick={() => choose(o)}
              >
                {words(shapeSample(o.shape!.sample).label)} — {o.anchor?.cell}
              </button>
              {o.protectedReason && <span>{o.protectedReason}</span>}
            </li>
          ))}
      </ul>
      {selected && (
        <button
          type="button"
          disabled={disabled || !selected.canDelete}
          onClick={() =>
            void apply({
              operation: 'deleteShape',
              sheet: frame.sheet,
              objectId: selected.id,
              expectedObjectRevision: selected.revision,
            })
          }
        >
          {words('Delete object')}
        </button>
      )}
    </section>
  );
}
