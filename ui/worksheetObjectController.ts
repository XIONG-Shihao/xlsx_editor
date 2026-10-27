/** A single physical read lane; original frames own reads and accepted commands. */
import type { WorkbookObjectCommand, WorkbookObjectPreview } from '../core/contracts';
import type { WorkbookObjectsProps } from './WorkbookObjects';
export interface WorksheetObjectFrame {
  key: string;
  family: string;
  sheet: string;
  active: boolean;
  reason?: string;
  controls: Pick<WorkbookObjectsProps, 'objects' | 'object' | 'onBatch'>;
}
export interface WorksheetObjectSnapshot {
  key: string;
  objects: WorkbookObjectPreview[];
  loading: boolean;
  pending: boolean;
  error: string;
}
export function createWorksheetObjectController() {
  let frame: WorksheetObjectFrame | null = null,
    wanted: WorksheetObjectFrame | null = null,
    reading = false,
    writing = false,
    disposed = false;
  let snapshot: WorksheetObjectSnapshot = { key: '', objects: [], loading: false, pending: false, error: '' };
  const listeners = new Set<() => void>(),
    emit = (next: WorksheetObjectSnapshot) => {
      snapshot = next;
      for (const listener of listeners) listener();
    };
  const same = (captured: WorksheetObjectFrame) => !disposed && frame?.key === captured.key && frame.active;
  async function drain() {
    if (reading || disposed) return;
    reading = true;
    try {
      while (wanted && !disposed) {
        const captured = wanted;
        wanted = null;
        if (!same(captured)) continue;
        emit({ key: captured.key, objects: [], loading: true, pending: writing, error: '' });
        try {
          const objects: WorkbookObjectPreview[] = [],
            ids = new Set<string>(),
            cursors = new Set<string>();
          let after: string | undefined,
            bytes = 0,
            count = 0;
          do {
            const page = await captured.controls.objects(captured.sheet, { after, limit: 32 });
            if (!same(captured)) break;
            if (
              !Array.isArray(page.objects) ||
              !Number.isSafeInteger(page.total) ||
              page.total < 0 ||
              page.total > 256 ||
              ++count > 8 ||
              page.objects.length > 32
            )
              throw Error('Worksheet object inventory exceeds its supported bound.');
            for (const entry of page.objects) {
              if (ids.has(entry.id)) throw Error('Worksheet object inventory repeats an identity.');
              ids.add(entry.id);
              if (!entry.anchor || !entry.managed) continue;
              if (objects.length >= 32) throw Error('Worksheet gestures support at most 32 owned objects.');
              const preview = await captured.controls.object(captured.sheet, entry.id);
              if (!same(captured)) break;
              if (
                preview.id !== entry.id ||
                preview.revision !== entry.revision ||
                JSON.stringify(preview.anchor) !== JSON.stringify(entry.anchor)
              )
                throw Error('Worksheet object preview changed its captured identity.');
              bytes += new TextEncoder().encode(JSON.stringify(preview)).byteLength;
              if (bytes > 8 * 1024 * 1024)
                throw Error('Worksheet object previews exceed their supported memory bound.');
              objects.push(preview);
            }
            if (!same(captured)) break;
            after = page.nextCursor ?? undefined;
            if (after) {
              if (cursors.has(after)) throw Error('Worksheet object inventory repeats a page.');
              cursors.add(after);
            }
          } while (after !== undefined);
          if (same(captured)) emit({ key: captured.key, objects, loading: false, pending: writing, error: '' });
        } catch (error) {
          if (same(captured))
            emit({
              key: captured.key,
              objects: [],
              loading: false,
              pending: writing,
              error: String(error instanceof Error ? error.message : error),
            });
        }
      }
    } finally {
      reading = false;
    }
  }
  return {
    subscribe: (listener: () => void) => {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    getSnapshot: () => snapshot,
    setFrame: (next: WorksheetObjectFrame | null) => {
      const changed = frame?.key !== next?.key || frame?.active !== next?.active;
      frame = next;
      if (!changed) return;
      wanted = next?.active ? { ...next, controls: { ...next.controls } } : null;
      emit({ key: next?.key ?? '', objects: [], loading: !!wanted, pending: writing, error: '' });
      void drain();
    },
    refresh: () => {
      if (frame?.active) {
        wanted = { ...frame, controls: { ...frame.controls } };
        void drain();
      }
    },
    accept: async (captured: WorksheetObjectFrame, command: WorkbookObjectCommand) => {
      if (!same(captured)) throw Error('The object or workbook changed. Select the object again.');
      if (frame?.reason || captured.reason) throw Error(frame?.reason ?? captured.reason);
      if (writing) throw Error('Wait for the object operation to finish.');
      command = structuredClone(command);
      const create = ['createShape', 'createSmartArt', 'createModel3D'].includes(command.operation);
      if (
        create &&
        (snapshot.key !== captured.key ||
          snapshot.loading ||
          snapshot.error ||
          snapshot.objects.length >= 32 ||
          command.expectedObjectRevision !== null ||
          snapshot.objects.some((item) => item.id === command.objectId))
      )
        throw Error('Wait for the current object inventory or remove an owned object before inserting.');
      const object =
        snapshot.key === captured.key ? snapshot.objects.find((item) => item.id === command.objectId) : undefined;
      if (
        command.sheet !== captured.sheet ||
        (!create &&
          (![
            'deleteObject',
            'updateImage',
            'updateChart',
            'updateShape',
            'deleteShape',
            'updateSmartArt',
            'deleteSmartArt',
            'updateModel3D',
            'deleteModel3D',
            'updateInkStroke',
            'deleteInkStroke',
          ].includes(command.operation) ||
            (command.operation === 'updateImage' && object?.kind !== 'image') ||
            (command.operation === 'updateChart' && object?.kind !== 'chart') ||
            (['updateShape', 'deleteShape'].includes(command.operation) && object?.kind !== 'shape') ||
            (['updateSmartArt', 'deleteSmartArt'].includes(command.operation) && object?.kind !== 'smartart') ||
            (['updateModel3D', 'deleteModel3D'].includes(command.operation) &&
              (object?.kind !== 'model3d' || !object.asset3d)) ||
            (['updateInkStroke', 'deleteInkStroke'].includes(command.operation) &&
              (object?.kind !== 'ink' || !object.stroke)) ||
            (object?.kind === 'model3d' && !object.asset3d) ||
            !object ||
            object.revision !== command.expectedObjectRevision ||
            object.protectedReason ||
            !object.managed ||
            (['deleteObject', 'deleteShape', 'deleteSmartArt', 'deleteModel3D', 'deleteInkStroke'].includes(
              command.operation
            )
              ? !object.canDelete
              : !object.canEdit)))
      )
        throw Error('The selected object is changed or protected.');
      writing = true;
      emit({ ...snapshot, pending: true, error: '' });
      try {
        await captured.controls.onBatch([command]);
        return { accepted: true, current: same(captured), key: captured.key };
      } catch (error) {
        if (same(captured)) emit({ ...snapshot, error: String(error instanceof Error ? error.message : error) });
        throw error;
      } finally {
        writing = false;
        emit({ ...snapshot, pending: false });
        if (wanted) void drain();
      }
    },
    dispose: () => {
      disposed = true;
      frame = null;
      wanted = null;
      listeners.clear();
    },
  };
}
export type WorksheetObjectController = ReturnType<typeof createWorksheetObjectController>;
