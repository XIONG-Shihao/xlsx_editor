import { type FormulaDraftHandle, type FormulaDraftLease, type FormulaDraftSnapshot } from '../namedFormulaDraft';
import { useLayoutEffect, useRef, useState } from 'react';

export function CellInput({
  value,
  display,
  pending,
  gridCell = false,
  onCommit,
  onMove,
  onBlur,
  onFocus,
  onKeyDown,
  onDoubleClick,
  readOnly,
  formulaBridge,
  previewBridge,
  ...props
}: Omit<React.InputHTMLAttributes<HTMLInputElement>, 'value'> & {
  value: string;
  display?: string;
  pending?: boolean;
  gridCell?: boolean;
  onCommit: (value: string) => void;
  onMove?: (key: 'Enter' | 'Tab', shift: boolean) => void;
  formulaBridge?: {
    handle: React.MutableRefObject<FormulaDraftHandle | null>;
    scopeKey: string;
    onComposition: (value: boolean) => void;
    open: () => void;
  };
  previewBridge?: {
    scopeKey: string;
    onDraft: (scopeKey: string, snapshot: FormulaDraftSnapshot, active: boolean) => void;
  };
}): JSX.Element {
  const [draft, setDraft] = useState(value);
  const [editing, setEditing] = useState(!gridCell);
  const focused = useRef(false);
  const last = useRef(value);
  const composing = useRef(false);
  const commitAfterComposition = useRef(false);
  const input = useRef<HTMLInputElement>(null),
    parked = useRef(false),
    lease = useRef<FormulaDraftLease | null>(null),
    draftRef = useRef(draft),
    generation = useRef(0),
    physicalBusy = useRef(false),
    pendingCaret = useRef<number | null>(null),
    pristineAfterScope = useRef(false);
  const previewActive = useRef(false);
  const notifyPreview = (active = previewActive.current): void =>
    previewBridge?.onDraft(
      previewBridge.scopeKey,
      {
        text: draftRef.current,
        start: input.current?.selectionStart ?? draftRef.current.length,
        end: input.current?.selectionEnd ?? draftRef.current.length,
        generation: generation.current,
        composing: composing.current,
      },
      active
    );
  draftRef.current = draft;
  const activeScope = useRef(formulaBridge?.scopeKey);
  activeScope.current = formulaBridge?.scopeKey;
  useLayoutEffect(() => {
    if (!formulaBridge) return;
    parked.current = false;
    lease.current = null;
    generation.current++;
    pristineAfterScope.current = true;
    composing.current = false;
    formulaBridge.onComposition(false);
    setDraft(value);
    draftRef.current = value;
    last.current = value;
  }, [formulaBridge?.scopeKey]);
  useLayoutEffect(() => {
    previewActive.current = false;
    if (previewBridge && !formulaBridge) {
      parked.current = false;
      generation.current++;
      pristineAfterScope.current = true;
      composing.current = false;
      setDraft(value);
      draftRef.current = value;
      last.current = value;
    }
    notifyPreview(false);
  }, [previewBridge?.scopeKey]);
  if (formulaBridge)
    formulaBridge.handle.current = {
      capture: () => ({
        text: draftRef.current,
        start: input.current?.selectionStart ?? draftRef.current.length,
        end: input.current?.selectionEnd ?? draftRef.current.length,
        generation: generation.current,
        composing: composing.current,
      }),
      park: () => {
        parked.current = true;
        previewActive.current = true;
        notifyPreview();
      },
      attach: (value) => {
        lease.current = value;
      },
      replace: (snapshot, text, caret) => {
        if (snapshot.generation !== generation.current || snapshot.text !== draftRef.current || composing.current)
          throw Error('The formula draft changed. Reopen Use in Formula.');
        generation.current++;
        pristineAfterScope.current = false;
        draftRef.current = text;
        pendingCaret.current = caret;
        setDraft(text);
        previewActive.current = true;
        notifyPreview();
      },
      focus: () => {
        input.current?.focus({ preventScroll: true });
      },
      ownsFocus: () => input.current === document.activeElement,
      isParked: () => parked.current,
    };
  useLayoutEffect(() => {
    if (pendingCaret.current !== null && input.current) {
      input.current.setSelectionRange(pendingCaret.current, pendingCaret.current);
      pendingCaret.current = null;
    }
  }, [draft]);
  useLayoutEffect(() => {
    // A new head masks its old viewport before the replacement tile arrives.
    // Adopt that tile even while focused, unless a new user draft has started.
    if (
      (!focused.current || !editing || ((formulaBridge || previewBridge) && pristineAfterScope.current)) &&
      !pending &&
      !parked.current &&
      !composing.current
    ) {
      setDraft(value);
      draftRef.current = value;
      last.current = value;
    }
  }, [value, pending]);
  const commit = (text: string): void => {
    if (lease.current) {
      const captured = lease.current,
        scope = activeScope.current;
      if (physicalBusy.current || !captured.isCurrent()) return;
      if (text === last.current) {
        parked.current = false;
        lease.current = null;
        return;
      }
      physicalBusy.current = true;
      void captured
        .accept(text)
        .then(() => {
          if (activeScope.current === scope) {
            last.current = text;
            parked.current = false;
            lease.current = null;
          }
        })
        .catch((error) => captured.error(error))
        .finally(() => {
          physicalBusy.current = false;
        });
      return;
    }
    if (text !== last.current) {
      last.current = text;
      onCommit(text);
    }
  };
  const description = [
    props['aria-description'],
    gridCell ? (editing ? 'Editing cell' : 'Press F2 to edit cell') : undefined,
  ]
    .filter(Boolean)
    .join('. ');
  return (
    <input
      {...props}
      ref={input}
      aria-description={description || undefined}
      data-editing={editing}
      readOnly={readOnly}
      value={gridCell && !editing ? (display ?? value) : draft}
      onFocus={(event) => {
        focused.current = true;
        previewActive.current = true;
        notifyPreview();
        onFocus?.(event);
      }}
      onChange={(event) => {
        if (gridCell && !editing) setEditing(true);
        generation.current++;
        pristineAfterScope.current = false;
        draftRef.current = event.target.value;
        setDraft(event.target.value);
        previewActive.current = true;
        notifyPreview();
      }}
      onDoubleClick={(event) => {
        if (gridCell && !readOnly) {
          setDraft(value);
          setEditing(true);
        }
        onDoubleClick?.(event);
      }}
      onCompositionStart={(event) => {
        composing.current = true;
        pristineAfterScope.current = false;
        formulaBridge?.onComposition(true);
        notifyPreview();
        event.currentTarget.dataset.composing = 'true';
        if (gridCell && !editing) {
          event.currentTarget.select();
          setDraft(event.currentTarget.value);
          setEditing(true);
        }
      }}
      onCompositionEnd={(event) => {
        composing.current = false;
        formulaBridge?.onComposition(false);
        event.currentTarget.dataset.composing = 'false';
        draftRef.current = event.currentTarget.value;
        setDraft(event.currentTarget.value);
        notifyPreview();
        if (commitAfterComposition.current) {
          commitAfterComposition.current = false;
          commit(event.currentTarget.value);
          if (gridCell) setEditing(false);
        }
      }}
      onBlur={(event) => {
        focused.current = false;
        if (
          (formulaBridge || previewBridge) &&
          (parked.current || (event.relatedTarget as HTMLElement | null)?.closest?.('[data-formula-draft-zone]'))
        ) {
          parked.current = true;
          onBlur?.(event);
          return;
        }
        previewActive.current = false;
        notifyPreview(false);
        if (editing && composing.current) commitAfterComposition.current = true;
        else if (editing) commit(event.currentTarget.value);
        if (gridCell && !composing.current) setEditing(false);
        onBlur?.(event);
      }}
      onKeyDown={(event) => {
        if (composing.current || event.nativeEvent.isComposing || event.keyCode === 229) return;
        if (formulaBridge && event.altKey && !event.ctrlKey && !event.metaKey && event.key === 'ArrowDown') {
          event.preventDefault();
          event.stopPropagation();
          formulaBridge.open();
          return;
        }
        if (
          gridCell &&
          !readOnly &&
          !editing &&
          (event.key === 'F2' || (event.key.length === 1 && !event.metaKey && !event.ctrlKey && !event.altKey))
        ) {
          event.preventDefault();
          event.stopPropagation();
          setEditing(true);
          setDraft(event.key === 'F2' ? value : event.key);
        } else if (editing && event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          parked.current = false;
          lease.current = null;
          generation.current++;
          draftRef.current = last.current;
          setDraft(last.current);
          previewActive.current = false;
          notifyPreview(false);
          if (gridCell) setEditing(false);
        } else if (editing && event.key === 'Enter') {
          event.preventDefault();
          event.stopPropagation();
          previewActive.current = false;
          notifyPreview(false);
          commit(draft);
          if (gridCell) {
            setEditing(false);
            onMove?.('Enter', event.shiftKey);
          } else if (!lease.current) event.currentTarget.blur();
        } else if (editing && event.key === 'Tab') {
          if (gridCell) {
            event.preventDefault();
            event.stopPropagation();
          }
          if (!formulaBridge) commit(draft);
          if (gridCell) {
            setEditing(false);
            onMove?.('Tab', event.shiftKey);
          }
        } else if (editing && event.key.startsWith('Arrow')) {
          event.stopPropagation();
        }
        onKeyDown?.(event);
      }}
    />
  );
}
