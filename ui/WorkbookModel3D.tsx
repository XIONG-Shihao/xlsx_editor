import { useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import type { WorkbookObjectCommand, WorkbookObjectPreview, WorkbookModel3DSample } from '../core/contracts';
import {
  MODEL3D_SAMPLES,
  normalizeModel3D,
  model3DGlb,
  parseModel3DGlb,
  model3DSampleAsset,
} from '../core/model3d-model.mjs';
import {
  createWorksheetObjectController,
  type WorksheetObjectFrame,
  type WorksheetObjectController,
} from './worksheetObjectController';
import { WorkbookModel3DPreview } from './WorkbookModel3DPreview';
const sampleAssets = MODEL3D_SAMPLES.map((sample) => ({ sample, asset: model3DSampleAsset(sample.id) }));
export function WorkbookModel3D({
  frame,
  selection,
  controller: sharedController,
  words = (s: string) => s,
}: {
  frame: WorksheetObjectFrame;
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
  const initial = () => ({
    key: frame.key,
    id: null as string | null,
    sample: 'cube' as WorkbookModel3DSample,
    color: '217346',
    rotation: { x: 20, y: -30, z: 0 },
    cell: selection.split(':')[0],
    width: 320,
    height: 240,
    dx: 0,
    dy: 0,
    alt: '',
  });
  const [draft, setDraft] = useState(initial),
    [notice, setNotice] = useState({ key: '', text: '' }),
    current = draft.key === frame.key ? draft : initial();
  const fresh = snapshot.key === frame.key,
    objects = fresh ? snapshot.objects : [],
    pending = snapshot.pending,
    selected = objects.find((o) => o.id === current.id),
    disabled = !frame.active || !!frame.reason || pending || !fresh || snapshot.loading || !!snapshot.error;
  const update = (patch: Partial<typeof draft>) => {
    if (live.current.key === frame.key && frame.active) setDraft({ ...current, ...patch });
  };
  const choose = (object: WorkbookObjectPreview) => {
    if (!object.asset3d || !object.anchor) return;
    const m = object.asset3d.model;
    update({
      id: object.id,
      sample: m.sample,
      color: m.color,
      rotation: { ...m.rotation },
      cell: object.anchor.cell,
      width: object.anchor.widthEmu / 9525,
      height: object.anchor.heightEmu / 9525,
      dx: object.anchor.dxEmu,
      dy: object.anchor.dyEmu,
      alt: object.altText ?? '',
    });
  };
  const model = () => normalizeModel3D({ sample: current.sample, color: current.color, rotation: current.rotation });
  const preview = useMemo(() => {
    try {
      return parseModel3DGlb(model3DGlb({ sample: current.sample, color: current.color, rotation: current.rotation }));
    } catch {
      return null;
    }
  }, [current.sample, current.color, current.rotation.x, current.rotation.y, current.rotation.z]);
  const reset = () => update(initial());
  async function apply(command: WorkbookObjectCommand) {
    if (live.current.key !== frame.key || !frame.active || disabled) return;
    try {
      setNotice({ key: frame.key, text: '' });
      const result = await controller.accept({ ...frame, controls: { ...frame.controls } }, command);
      if (result.current) {
        setDraft(initial());
        controller.refresh();
      }
    } catch (error) {
      if (live.current.key === frame.key)
        setNotice({ key: frame.key, text: String(error instanceof Error ? error.message : error) });
    }
  }
  return (
    <section aria-label={words('3D Models')} style={{ padding: 12, display: 'grid', gap: 12 }}>
      <p>
        {words(
          'Choose a solid model, color and rotation. Select a saved model below to edit it; move and resize it on the worksheet.'
        )}
      </p>
      <div
        role="group"
        aria-label={words('3D model samples')}
        style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(110px,1fr))', gap: 8 }}
      >
        {sampleAssets.map(({ sample, asset }) => (
          <button
            type="button"
            key={sample.id}
            aria-label={words(sample.label)}
            aria-pressed={current.sample === sample.id}
            disabled={disabled}
            onClick={() => update({ sample: sample.id })}
            style={{ padding: 8 }}
          >
            <span style={{ display: 'block', height: 70, pointerEvents: 'none' }}>
              <WorkbookModel3DPreview asset={asset} width={180} height={108} />
            </span>
            {words(sample.label)}
          </button>
        ))}
      </div>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          try {
            const model3d = model(),
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
                    operation: 'updateModel3D',
                    sheet: frame.sheet,
                    objectId: selected.id,
                    expectedObjectRevision: selected.revision,
                    patch: model3d,
                    anchor,
                    altText: current.alt,
                  }
                : {
                    operation: 'createModel3D',
                    sheet: frame.sheet,
                    objectId: crypto.randomUUID(),
                    expectedObjectRevision: null,
                    model3d,
                    anchor,
                    altText: current.alt,
                  }
            );
          } catch (error) {
            if (live.current.key === frame.key && frame.active)
              setNotice({ key: frame.key, text: String(error instanceof Error ? error.message : error) });
          }
        }}
      >
        <fieldset disabled={disabled || (!!selected && !selected.canEdit)} style={{ display: 'grid', gap: 10 }}>
          <legend>{words(selected ? 'Edit 3D model' : 'Insert 3D model')}</legend>
          <label>
            {words('Model color')}
            <input
              aria-label={words('3D model color')}
              type="color"
              value={'#' + current.color}
              onChange={(e) => update({ color: e.target.value.slice(1).toUpperCase() })}
            />
          </label>
          {(['x', 'y', 'z'] as const).map((axis) => (
            <label key={axis}>
              {words('Rotation') + ' ' + axis.toUpperCase() + ' (°)'}
              <input
                aria-label={words('3D rotation') + ' ' + axis.toUpperCase()}
                type="number"
                min="-180"
                max="180"
                step="1"
                required
                value={Number.isFinite(current.rotation[axis]) ? current.rotation[axis] : ''}
                onChange={(e) => update({ rotation: { ...current.rotation, [axis]: Number(e.target.value) } })}
              />
            </label>
          ))}
          <label>
            {words('Anchor cell')}
            <input
              aria-label={words('Object anchor')}
              required
              value={current.cell}
              onChange={(e) => update({ cell: e.target.value, dx: 0, dy: 0 })}
            />
          </label>
          <label>
            {words('Width in pixels')}
            <input
              aria-label={words('Object width')}
              type="number"
              min={1 / 9525}
              max={50000000 / 9525}
              step="any"
              required
              value={Number.isFinite(current.width) ? current.width : ''}
              onChange={(e) => update({ width: Number(e.target.value) })}
            />
          </label>
          <label>
            {words('Height in pixels')}
            <input
              aria-label={words('Object height')}
              type="number"
              min={1 / 9525}
              max={50000000 / 9525}
              step="any"
              required
              value={Number.isFinite(current.height) ? current.height : ''}
              onChange={(e) => update({ height: Number(e.target.value) })}
            />
          </label>
          <label>
            {words('Alternative text')}
            <input
              aria-label={words('Object alternative text')}
              maxLength={256}
              value={current.alt}
              onChange={(e) => update({ alt: e.target.value })}
            />
          </label>
          <p>
            {words(
              'Cube, Pyramid and Octahedron use solid colors and whole-degree rotations from −180° to 180°. Imported 3D models and textures are not supported.'
            )}
          </p>
          <div style={{ height: 180, pointerEvents: 'none' }}>
            {preview ? (
              <WorkbookModel3DPreview
                asset={preview}
                width={
                  Number.isFinite(current.width) && current.width >= 1 / 9525 && current.width <= 50000000 / 9525
                    ? current.width
                    : 320
                }
                height={
                  Number.isFinite(current.height) && current.height >= 1 / 9525 && current.height <= 50000000 / 9525
                    ? current.height
                    : 240
                }
                label={current.alt}
              />
            ) : (
              <p role="status">{words('Choose a valid color and rotation to preview the model.')}</p>
            )}
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
          .filter((o) => o.asset3d)
          .map((o) => (
            <li key={o.id}>
              <button
                type="button"
                data-model3d-object={o.id}
                aria-pressed={selected?.id === o.id}
                disabled={pending || !fresh || snapshot.loading || !!snapshot.error || !frame.active}
                onClick={() => choose(o)}
              >
                {words(MODEL3D_SAMPLES.find((s) => s.id === o.asset3d!.model.sample)!.label)} — {o.anchor?.cell}
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
              operation: 'deleteModel3D',
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
