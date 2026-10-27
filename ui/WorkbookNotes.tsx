import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import { createNoteController, type NoteFrame } from './noteController';
import { NOTE_LIMITS } from '../core/notes.mjs';
export function WorkbookNotes({
  frame,
  words = (s: string) => s,
}: {
  frame: NoteFrame | null;
  words?: (s: string) => string;
}) {
  const [controller] = useState(createNoteController),
    lifetime = useRef(0);
  const snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot);
  useLayoutEffect(() => controller.setFrame(frame));
  useEffect(() => {
    const serial = ++lifetime.current;
    return () => {
      controller.setFrame(null);
      queueMicrotask(() => {
        if (lifetime.current === serial) controller.dispose();
      });
    };
  }, [controller]);
  const inventory = snapshot.key === frame?.key ? snapshot.inventory : null;
  const selected = inventory?.notes?.find((n) => n.cell === frame?.cell);
  const [draft, setDraft] = useState({ key: '', author: 'Author', text: '' }),
    [error, setError] = useState({ key: '', text: '' });
  useEffect(() => {
    if (!frame || !inventory) return;
    setDraft({ key: frame.key, author: selected?.author ?? 'Author', text: selected?.text ?? '' });
    setError({ key: '', text: '' });
  }, [inventory, frame?.key]);
  const current = draft.key === frame?.key ? draft : { key: frame?.key ?? '', author: 'Author', text: '' };
  const disabled = !frame?.active || !frame.writable || !inventory?.canEdit || snapshot.loading || snapshot.pending;
  const edit = (patch: Partial<typeof draft>) => {
    if (frame?.active && !disabled) {
      setDraft({ ...current, ...patch });
      setError({ key: '', text: '' });
    }
  };
  const run = async (action: () => Promise<unknown> | void) => {
    const key = frame?.key;
    try {
      setError({ key: '', text: '' });
      await action();
    } catch (cause) {
      if (key === frame?.key)
        setError({ key: key ?? '', text: String(cause instanceof Error ? cause.message : cause) });
    }
  };
  return (
    <section aria-label={words('Worksheet notes')} style={{ display: 'grid', gap: 12, padding: 12 }}>
      <p>
        {words('Selected cell')}: {frame ? frame.sheet + '!' + frame.cell : words('Select a regular cell.')}
      </p>
      <p>{words('Notes keep the cell contents and formatting. Choose a note location to view or edit it.')}</p>
      {frame?.reason && <p role="status">{words(frame.reason)}</p>}
      {inventory?.protectedReason && <p role="status">{words(inventory.protectedReason)}</p>}
      {snapshot.loading && <p role="status">{words('Loading worksheet notes…')}</p>}
      {inventory && (
        <div>
          <h3>{words('Notes on this worksheet')}</h3>
          {inventory.notes === null ? (
            <p>
              {words(
                'This imported note graph is protected. Its contents are preserved; an editable list is unavailable.'
              )}
            </p>
          ) : inventory.notes.length ? (
            <ul>
              {inventory.notes.map((note) => (
                <li key={note.cell}>
                  <button
                    type="button"
                    disabled={!frame?.active || snapshot.pending || snapshot.loading}
                    aria-label={words('Go to note') + ' ' + note.cell}
                    onClick={() => {
                      if (frame) void run(() => controller.navigate(frame, note.cell));
                    }}
                  >
                    {note.cell + ' — ' + note.author}
                  </button>
                </li>
              ))}
            </ul>
          ) : (
            <p>{words('No notes on this worksheet.')}</p>
          )}
        </div>
      )}
      <form
        onSubmit={(e) => {
          e.preventDefault();
          if (frame && !disabled)
            void run(() => controller.write(frame, { author: current.author, text: current.text }));
        }}
      >
        <fieldset disabled={disabled} style={{ display: 'grid', gap: 10 }}>
          <legend>{words(selected ? 'Edit note' : 'New note')}</legend>
          <label>
            {words('Note author')}
            <input
              aria-label={words('Note author')}
              value={current.author}
              required
              maxLength={NOTE_LIMITS.author}
              onChange={(e) => edit({ author: e.target.value })}
            />
          </label>
          <label>
            {words('Note text')}
            <textarea
              aria-label={words('Note text')}
              value={current.text}
              required
              rows={6}
              maxLength={NOTE_LIMITS.text}
              onChange={(e) => edit({ text: e.target.value })}
            />
          </label>
          <button type="submit">{words(selected ? 'Apply note changes' : 'Add note')}</button>
        </fieldset>
      </form>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
        <button
          type="button"
          disabled={disabled || !selected}
          onClick={() => {
            if (frame) void run(() => controller.write(frame, null));
          }}
        >
          {words('Delete note')}
        </button>
        <button
          type="button"
          disabled={!frame?.active || snapshot.loading || snapshot.pending}
          onClick={() => {
            setError({ key: '', text: '' });
            controller.refresh();
          }}
        >
          {words('Reset note form')}
        </button>
      </div>
      {snapshot.pending && <p role="status">{words('Applying note change…')}</p>}
      {((snapshot.key === frame?.key && snapshot.error) || (error.key === frame?.key && error.text)) && (
        <p role="alert">{words((snapshot.key === frame?.key && snapshot.error) || error.text)}</p>
      )}
    </section>
  );
}
