import { useEffect, useState } from 'react';

export function DimensionInput({
  value,
  min,
  max,
  disabled,
  onCommit,
  onError,
  commandId,
  onDraft,
}: {
  value: number;
  min: number;
  max: number;
  disabled?: boolean;
  commandId?: 'home.column-width' | 'home.row-height';
  onDraft?: (value: string, accept: () => void) => void;
  onCommit: (value: number) => void;
  onError: (message: string) => void;
}): JSX.Element {
  const [draft, setDraft] = useState(String(value));
  useEffect(() => setDraft(String(value)), [value]);
  return (
    <input
      data-home-command={commandId}
      data-home-dimension
      type="number"
      min={min}
      max={max}
      step="any"
      value={draft}
      disabled={disabled}
      onChange={(event) => {
        const value = event.target.value;
        if (onDraft) onDraft(value, () => setDraft(value));
        else setDraft(value);
      }}
      onKeyDown={(event) => {
        if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          setDraft(String(value));
          onError('');
        }
        if (event.key === 'Enter') {
          event.preventDefault();
          event.currentTarget.blur();
        }
      }}
      onBlur={() => {
        if (draft.trim() !== '' && Number(draft) === value) return;
        const next = Number(draft);
        if (!draft.trim() || !Number.isFinite(next) || next < min || next > max) {
          setDraft(String(value));
          onError(`Enter a size between ${min} and ${max}.`);
          return;
        }
        onError('');
        onCommit(next);
      }}
    />
  );
}
