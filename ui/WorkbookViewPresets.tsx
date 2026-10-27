import { useLayoutEffect, useRef, useState } from 'react';
import { BUILTIN_VIEW_PRESETS, type NamedViewPreset } from './viewPresets';
import type { ViewPreferences } from './viewPreferences';
import type { RibbonCommandId } from './ribbonCommands';
export interface ViewPresetControls {
  key: string;
  active: boolean;
  preferences: ViewPreferences;
  named: NamedViewPreset[];
  warning: string | null;
  apply(value: ViewPreferences): void;
  save(name: string): void;
  update(name: string): void;
  remove(name: string): void;
  reset(): void;
}
export function WorkbookViewPresets({
  controls,
  run,
  words,
}: {
  controls: ViewPresetControls;
  run(id: RibbonCommandId, handler: () => void): Promise<void>;
  words(text: string): string;
}): JSX.Element {
  const [selected, setSelected] = useState(''),
    [label, setLabel] = useState(''),
    [error, setError] = useState(''),
    live = useRef(controls),
    mounted = useRef(true);
  live.current = controls;
  useLayoutEffect(() => {
    setSelected('');
    setLabel('');
    setError('');
  }, [controls.key]);
  useLayoutEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
    };
  }, []);
  const current = () => mounted.current && live.current === controls && live.current.active;
  const named = selected.startsWith('named:')
    ? controls.named.find((item) => item.name === selected.slice(6))
    : undefined;
  const invoke = async (id: RibbonCommandId, action: () => void): Promise<void> => {
    if (!current()) return;
    setError('');
    await run(id, () => {
      if (!current()) return;
      try {
        action();
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    });
  };
  return (
    <div className="sheet-ribbon-stack" data-formula-draft-zone>
      <label>
        {words('Local view')}
        <select
          aria-label={words('Local view preset')}
          data-ribbon-command="view.preset-apply"
          value={named || selected.startsWith('builtin:') ? selected : ''}
          disabled={!controls.active}
          onChange={(event) => {
            if (!current()) return;
            const id = event.target.value;
            setSelected(id);
            const value = id.startsWith('builtin:')
              ? BUILTIN_VIEW_PRESETS.find((p) => p.id === id.slice(8))?.preferences
              : controls.named.find((p) => p.name === id.slice(6))?.preferences;
            if (value) void invoke('view.preset-apply', () => controls.apply({ ...value }));
          }}
        >
          <option value="">{words('Current view')}</option>
          <optgroup label={words('Built-in views')}>
            {BUILTIN_VIEW_PRESETS.map((p) => (
              <option key={p.id} value={'builtin:' + p.id}>
                {words(p.name)}
              </option>
            ))}
          </optgroup>
          {controls.named.length > 0 && (
            <optgroup label={words('Saved for this sheet')}>
              {controls.named.map((p) => (
                <option key={p.name} value={'named:' + p.name}>
                  {p.name}
                </option>
              ))}
            </optgroup>
          )}
        </select>
      </label>
      <label>
        {words('Save current view as')}
        <input
          aria-label={words('Local view name')}
          maxLength={64}
          value={label}
          disabled={!controls.active}
          onChange={(event) => {
            if (current()) setLabel(event.target.value);
          }}
        />
      </label>
      <div>
        <button
          type="button"
          data-ribbon-command="view.preset-save"
          disabled={!controls.active || !label.trim()}
          onClick={() =>
            void invoke('view.preset-save', () => {
              controls.save(label);
              setSelected('named:' + label.trim());
              setLabel('');
            })
          }
        >
          {words('Save local view')}
        </button>
        <button
          type="button"
          data-ribbon-command="view.preset-update"
          disabled={!controls.active || !named}
          onClick={() => void invoke('view.preset-update', () => controls.update(named!.name))}
        >
          {words('Update saved view')}
        </button>
        <button
          type="button"
          data-ribbon-command="view.preset-delete"
          disabled={!controls.active || !named}
          onClick={() =>
            void invoke('view.preset-delete', () => {
              controls.remove(named!.name);
              setSelected('');
            })
          }
        >
          {words('Delete saved view')}
        </button>
      </div>
      <button
        type="button"
        data-ribbon-command="view.source-defaults"
        disabled={!controls.active}
        onClick={() =>
          void invoke('view.source-defaults', () => {
            controls.reset();
            setSelected('');
          })
        }
      >
        {words('Use workbook view')}
      </button>
      {error && <span role="alert">{words(error)}</span>}
      {controls.warning && <span role="status">{words(controls.warning)}</span>}
    </div>
  );
}
