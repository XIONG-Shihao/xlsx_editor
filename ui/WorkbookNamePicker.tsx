import { useLayoutEffect, useRef, useState } from 'react';
import type { NamedFormulaChoice } from './namedFormulaDraft';
export function WorkbookNamePicker({
  choices,
  target,
  onChoose,
  onClose,
  words = (text: string) => text,
}: {
  choices: readonly NamedFormulaChoice[];
  target: string;
  onChoose: (choice: NamedFormulaChoice) => void;
  onClose: () => void;
  words?: (text: string) => string;
}): JSX.Element {
  const [query, setQuery] = useState(''),
    [page, setPage] = useState(0),
    search = useRef<HTMLInputElement>(null),
    root = useRef<HTMLElement>(null),
    [position, setPosition] = useState({ left: 8, top: 8 });
  useLayoutEffect(() => {
    const box = root.current?.parentElement?.getBoundingClientRect();
    if (box)
      setPosition({
        left: Math.max(8, Math.min(box.left, window.innerWidth - 368)),
        top: Math.max(8, Math.min(box.bottom + 4, window.innerHeight - 360)),
      });
    search.current?.focus({ preventScroll: true });
  }, []);
  const found = choices.filter((choice) =>
      (choice.name + ' ' + choice.scope).toLowerCase().includes(query.toLowerCase())
    ),
    items = found.slice(page * 100, (page + 1) * 100);
  return (
    <section
      ref={root}
      role="dialog"
      aria-label={words('Use in Formula')}
      className="sheet-name-picker"
      style={position}
      data-formula-draft-zone
      onKeyDown={(event) => {
        if (event.key === 'Escape' && !event.nativeEvent.isComposing) {
          event.preventDefault();
          event.stopPropagation();
          onClose();
        }
      }}
    >
      <header>
        <strong>{words('Use in Formula')}</strong>
        <button type="button" aria-label={words('Close name picker')} onClick={onClose}>
          ×
        </button>
      </header>
      <p>
        {words('Insert a name in the formula draft for')} <strong>{target}</strong>.{' '}
        {words('Press Enter in the formula bar to apply it.')}
      </p>
      <input
        ref={search}
        aria-label={words('Search names')}
        value={query}
        onChange={(event) => {
          setQuery(event.target.value);
          setPage(0);
        }}
      />
      <ul aria-label={words('Available formula names')}>
        {items.map((choice) => (
          <li key={choice.id}>
            <button
              type="button"
              disabled={!!choice.disabledReason}
              title={words(choice.disabledReason ?? choice.formula)}
              aria-label={words('Insert name') + ' ' + choice.name + ' · ' + choice.scope}
              onClick={() => onChoose(choice)}
            >
              <strong>{choice.name}</strong>
              <small>{choice.scope}</small>
              <span>{choice.formula}</span>
            </button>
            {choice.disabledReason && <small>{words(choice.disabledReason)}</small>}
          </li>
        ))}
      </ul>
      {!found.length && <p role="status">{words('No matching names.')}</p>}
      {found.length > 100 && (
        <nav>
          <button type="button" disabled={!page} onClick={() => setPage((value) => value - 1)}>
            {words('Previous names')}
          </button>
          <span>
            {page * 100 + 1}–{Math.min((page + 1) * 100, found.length)} / {found.length}
          </span>
          <button
            type="button"
            disabled={(page + 1) * 100 >= found.length}
            onClick={() => setPage((value) => value + 1)}
          >
            {words('More names')}
          </button>
        </nav>
      )}
      <button type="button" onClick={onClose}>
        {words('Cancel name insertion')}
      </button>
    </section>
  );
}
