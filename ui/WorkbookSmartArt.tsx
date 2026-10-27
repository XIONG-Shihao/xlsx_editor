import { useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import type { WorkbookObjectCommand, WorkbookObjectPreview, WorkbookSmartArtLayout } from '../core/contracts';
import { SMARTART_SAMPLES, normalizeSmartArt, smartartSampleModel } from '../core/smartart-model.mjs';
import {
  createWorksheetObjectController,
  type WorksheetObjectFrame,
  type WorksheetObjectController,
} from './worksheetObjectController';
import { WorkbookSmartArtPreview } from './WorkbookSmartArtPreview';
export function WorkbookSmartArt({
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
    layout: 'process' as WorkbookSmartArtLayout,
    labels: smartartSampleModel('process').nodes.map((n) => n.text),
    touched: false,
    cell: selection.split(':')[0],
    width: 600,
    height: 360,
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
    if (!object.smartart || !object.anchor) return;
    update({
      id: object.id,
      layout: object.smartart.layout,
      labels: object.smartart.nodes.map((n) => n.text),
      touched: true,
      cell: object.anchor.cell,
      width: object.anchor.widthEmu / 9525,
      height: object.anchor.heightEmu / 9525,
      dx: object.anchor.dxEmu,
      dy: object.anchor.dyEmu,
      alt: object.altText ?? '',
    });
  };
  const model = () =>
    normalizeSmartArt({
      layout: current.layout,
      nodes: current.labels.map((text, i) => ({ id: 'n' + (i + 1), text })),
    });
  let preview: ReturnType<typeof model> | null = null;
  try {
    preview = model();
  } catch {}
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
    <section aria-label={words('SmartArt')} style={{ padding: 12, display: 'grid', gap: 12 }}>
      <p>
        {words(
          'Choose a layout and enter two to six node labels. Select a saved diagram below to edit it; move and resize it on the worksheet.'
        )}
      </p>
      <div
        role="group"
        aria-label={words('SmartArt layouts')}
        style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(110px,1fr))', gap: 8 }}
      >
        {SMARTART_SAMPLES.map((sample) => (
          <button
            type="button"
            key={sample.id}
            aria-label={words(sample.label)}
            aria-pressed={current.layout === sample.id}
            disabled={disabled}
            onClick={() =>
              update({
                layout: sample.id,
                labels:
                  current.touched || current.id
                    ? current.labels
                    : smartartSampleModel(sample.id).nodes.map((n) => n.text),
              })
            }
            style={{ padding: 8 }}
          >
            <span style={{ display: 'block', height: 70, pointerEvents: 'none' }}>
              <WorkbookSmartArtPreview model={smartartSampleModel(sample.id)} width={180} height={108} />
            </span>
            {words(sample.label)}
          </button>
        ))}
      </div>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          try {
            const smartart = model(),
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
                    operation: 'updateSmartArt',
                    sheet: frame.sheet,
                    objectId: selected.id,
                    expectedObjectRevision: selected.revision,
                    patch: smartart,
                    anchor,
                    altText: current.alt,
                  }
                : {
                    operation: 'createSmartArt',
                    sheet: frame.sheet,
                    objectId: crypto.randomUUID(),
                    expectedObjectRevision: null,
                    smartart,
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
          <legend>{words(selected ? 'Edit SmartArt' : 'Insert SmartArt')}</legend>
          <label>
            {words('Node count')}
            <select
              aria-label={words('SmartArt node count')}
              value={current.labels.length}
              onChange={(e) => {
                const count = Number(e.target.value);
                update({
                  touched: true,
                  labels: Array.from({ length: count }, (_, i) => current.labels[i] ?? 'Node ' + (i + 1)),
                });
              }}
            >
              {[2, 3, 4, 5, 6].map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
          </label>
          {current.labels.map((text, i) => (
            <label key={i}>
              {words('Node') + ' ' + (i + 1)}
              <input
                aria-label={words('SmartArt node') + ' ' + (i + 1)}
                maxLength={128}
                required
                value={text}
                onChange={(e) =>
                  update({ touched: true, labels: current.labels.map((v, j) => (j === i ? e.target.value : v)) })
                }
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
              min="1"
              max="5249"
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
              min="1"
              max="5249"
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
              'Process, Hierarchy and Cycle use fixed blue/white colors and Arial 14 pt. Each node supports one nonblank paragraph up to 128 characters.'
            )}
          </p>
          <div style={{ height: 180, pointerEvents: 'none' }}>
            {preview ? (
              <WorkbookSmartArtPreview
                model={preview}
                width={Number.isFinite(current.width) && current.width > 0 ? current.width : 600}
                height={Number.isFinite(current.height) && current.height > 0 ? current.height : 360}
                label={current.alt}
              />
            ) : (
              <p role="status">{words('Enter valid nonblank labels to preview the diagram.')}</p>
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
          .filter((o) => o.smartart)
          .map((o) => (
            <li key={o.id}>
              <button
                type="button"
                data-smartart-object={o.id}
                aria-pressed={selected?.id === o.id}
                disabled={pending || !fresh || snapshot.loading || !!snapshot.error || !frame.active}
                onClick={() => choose(o)}
              >
                {words(SMARTART_SAMPLES.find((s) => s.id === o.smartart!.layout)!.label)} — {o.anchor?.cell}
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
              operation: 'deleteSmartArt',
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
