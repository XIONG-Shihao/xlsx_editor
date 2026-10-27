import { useLayoutEffect, useRef, useState } from 'react';
import {
  LOCAL_ACTION_CATALOG,
  LOCAL_ACTION_LIMITS,
  localActionCatalog,
  validateLocalActionDraft,
  validateLocalActionList,
  type LocalActionDescriptor,
  type LocalActionDraft,
  type LocalActionEntry,
  type LocalActionFormat,
  type LocalActionList,
} from '../core/local-actions.mjs';

export interface WorkbookActionsControls {
  list: () => Promise<LocalActionList>;
  create: (draft: LocalActionDraft, expectedRegistryRevision: number) => Promise<LocalActionList>;
  update: (id: string, draft: LocalActionDraft, expectedRegistryRevision: number) => Promise<LocalActionList>;
  remove: (id: string, expectedRegistryRevision: number) => Promise<LocalActionList>;
}
export interface WorkbookActionsFrame {
  key: string;
  sheet?: string;
  controls?: WorkbookActionsControls;
}
export interface WorkbookActionsProps {
  frame: WorkbookActionsFrame;
  visible?: boolean;
  openIntent?: { serial: number; mode: 'new' | 'list' };
  words?: (text: string) => string;
}
type Command = LocalActionDescriptor['command'];
type TriState = '' | 'true' | 'false';
interface Form {
  name: string;
  command: Command;
  sheet: string;
  rows: string;
  columns: string;
  range: string;
  bold: TriState;
  italic: TriState;
  numberFormat: string;
  otherFormat: LocalActionFormat;
  keys: Array<{ column: string; direction: 'asc' | 'desc' }>;
}
const empty = (sheet = 'Sheet1'): Form => ({
  name: '',
  command: 'freeze',
  sheet,
  rows: '1',
  columns: '0',
  range: '',
  bold: '',
  italic: '',
  numberFormat: '',
  otherFormat: {},
  keys: [{ column: 'A', direction: 'asc' }],
});
const tri = (value: boolean | undefined): TriState => (value === undefined ? '' : value ? 'true' : 'false');
function fromEntry(entry: LocalActionEntry): Form {
  const result = { ...empty(), name: entry.name, command: entry.descriptor.command },
    descriptor = entry.descriptor;
  if (descriptor.command === 'recalculate') return result;
  result.sheet = descriptor.parameters.sheet;
  if (descriptor.command === 'freeze') {
    result.rows = String(descriptor.parameters.rows);
    result.columns = String(descriptor.parameters.columns);
  }
  if (descriptor.command === 'sortRows') {
    result.range = descriptor.parameters.range;
    result.keys = descriptor.parameters.keys.map((key) => ({ ...key }));
  }
  if (descriptor.command === 'formatRange') {
    result.range = descriptor.parameters.range;
    const { bold, italic, numberFormat, ...rest } = descriptor.parameters.format;
    result.bold = tri(bold);
    result.italic = tri(italic);
    result.numberFormat = numberFormat ?? '';
    result.otherFormat = rest;
  }
  return result;
}
function draft(form: Form): LocalActionDraft {
  let descriptor: LocalActionDescriptor;
  if (form.command === 'freeze')
    descriptor = {
      schema: 1,
      command: 'freeze',
      parameters: {
        sheet: form.sheet,
        rows: /^\d+$/.test(form.rows) ? Number(form.rows) : NaN,
        columns: /^\d+$/.test(form.columns) ? Number(form.columns) : NaN,
      },
    };
  else if (form.command === 'recalculate') descriptor = { schema: 1, command: 'recalculate', parameters: {} };
  else if (form.command === 'sortRows')
    descriptor = {
      schema: 1,
      command: 'sortRows',
      parameters: { sheet: form.sheet, range: form.range, keys: form.keys.map((key) => ({ ...key })) },
    };
  else
    descriptor = {
      schema: 1,
      command: 'formatRange',
      parameters: {
        sheet: form.sheet,
        range: form.range,
        format: {
          ...form.otherFormat,
          ...(form.bold === '' ? {} : { bold: form.bold === 'true' }),
          ...(form.italic === '' ? {} : { italic: form.italic === 'true' }),
          ...(form.numberFormat ? { numberFormat: form.numberFormat as LocalActionFormat['numberFormat'] } : {}),
        },
      },
    };
  return validateLocalActionDraft({ name: form.name, descriptor });
}
const identity = (text: string) => text;
const equivalent = (a: unknown, b: unknown): boolean => {
  const stable = (value: unknown): string =>
    JSON.stringify(value, (_key, item) =>
      item && typeof item === 'object' && !Array.isArray(item)
        ? Object.fromEntries(
            Object.keys(item)
              .sort()
              .map((key) => [key, item[key]])
          )
        : item
    );
  return stable(a) === stable(b);
};
const formats = [
  'general',
  'number',
  'percent',
  'date',
  'text',
  'currency',
  'accounting',
  'thousands',
  'decimal0',
  'decimal1',
  'decimal2',
] as const;
const numberLabels: Record<string, string> = {
  general: 'General',
  number: 'Number',
  percent: 'Percent',
  date: 'Date',
  text: 'Text',
  currency: 'Currency',
  accounting: 'Accounting',
  thousands: 'Thousands separator',
  decimal0: 'No decimals',
  decimal1: 'One decimal',
  decimal2: 'Two decimals',
};
const formatLabels: Record<keyof LocalActionFormat, string> = {
  bold: 'Bold',
  italic: 'Italic',
  underline: 'Underline',
  fontFamily: 'Font family',
  fontSize: 'Font size',
  fontColor: 'Font color',
  fillColor: 'Fill color',
  horizontalAlign: 'Horizontal alignment',
  verticalAlign: 'Vertical alignment',
  wrapText: 'Wrap text',
  border: 'Borders',
  numberFormat: 'Number format',
};

/** A private definition editor. It has no workbook edit, script or Run callback. */
export function WorkbookActions({
  frame,
  visible = true,
  openIntent,
  words = identity,
}: WorkbookActionsProps): JSX.Element {
  const key = frame.key,
    defaultSheet = frame.sheet ?? 'Sheet1',
    available = !!frame.controls && typeof key === 'string' && key.length > 0;
  const live = useRef({ key, visible, available });
  live.current = { key, visible, available };
  const mounted = useRef(false),
    generation = useRef(0),
    physical = useRef<{ kind: 'loading' | 'saving' | 'deleting' } | null>(null),
    loadWanted = useRef(false),
    lastIntent = useRef(-1);
  const [list, setList] = useState<{ key: string; value: LocalActionList } | null>(null),
    [form, setForm] = useState<{ key: string; value: Form }>({ key, value: empty(defaultSheet) }),
    [mode, setMode] = useState<'list' | 'new' | 'edit'>('list'),
    [editing, setEditing] = useState<{ id: string; revision: number } | null>(null),
    [phase, setPhase] = useState<'idle' | 'loading' | 'saving' | 'deleting'>('idle'),
    [error, setError] = useState(''),
    [notice, setNotice] = useState(''),
    [loadTick, setLoadTick] = useState(0);
  const current = () => mounted.current && live.current.key === key && live.current.visible && live.current.available;
  const fresh = list?.key === key ? list.value : null,
    fields = form.key === key ? form.value : empty(),
    busy = phase !== 'idle',
    selected = editing && fresh?.entries.find((entry) => entry.id === editing.id),
    editCurrent = mode !== 'edit' || (!!selected && selected.revision === editing?.revision);
  const clearMessages = () => {
    setError('');
    setNotice('');
  };
  function finish(job: { kind: 'loading' | 'saving' | 'deleting' }) {
    if (physical.current !== job) return;
    physical.current = null;
    if (mounted.current) {
      setPhase('idle');
      if (loadWanted.current && live.current.visible && live.current.available) setLoadTick((value) => value + 1);
    }
  }
  async function load(): Promise<void> {
    if (!current() || !frame.controls) return;
    if (physical.current) {
      loadWanted.current = true;
      return;
    }
    loadWanted.current = false;
    const job = { kind: 'loading' as const },
      token = ++generation.current,
      controls = frame.controls;
    physical.current = job;
    setPhase('loading');
    clearMessages();
    try {
      const result = validateLocalActionList(await controls.list());
      if (current() && generation.current === token) {
        if (fresh && result.revision < fresh.revision)
          throw Error('The library reply is older than the displayed registry.');
        setList({ key, value: result });
      }
    } catch (failure) {
      if (current() && generation.current === token)
        setError(failure instanceof Error ? failure.message : 'The private action library could not be loaded.');
    } finally {
      finish(job);
    }
  }
  useLayoutEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      generation.current++;
    };
  }, []);
  useLayoutEffect(() => {
    generation.current++;
    setList(null);
    setForm({ key, value: empty(defaultSheet) });
    setEditing(null);
    setMode('list');
    clearMessages();
    setPhase(physical.current?.kind ?? 'idle');
  }, [key]);
  useLayoutEffect(() => {
    generation.current++;
    if (visible && available) {
      loadWanted.current = true;
      void load();
    } else loadWanted.current = false;
    return () => {
      generation.current++;
    };
  }, [key, visible, available, loadTick]);
  useLayoutEffect(() => {
    if (!openIntent) return;
    if (
      Object.keys(openIntent).sort().join(',') !== 'mode,serial' ||
      !Number.isSafeInteger(openIntent.serial) ||
      openIntent.serial < 0 ||
      !['new', 'list'].includes(openIntent.mode)
    )
      throw Error('Invalid Local Actions open intent.');
    if (openIntent.serial <= lastIntent.current || !current()) return;
    lastIntent.current = openIntent.serial;
    if (physical.current && physical.current.kind !== 'loading') return;
    if (openIntent.mode === 'new' && mode !== 'new') {
      setForm({ key, value: empty(defaultSheet) });
      setEditing(null);
    }
    setMode(openIntent.mode);
    clearMessages();
  }, [openIntent?.serial, openIntent?.mode, key, visible]);
  const editForm = (patch: Partial<Form>) => {
    if (!current() || physical.current?.kind === 'saving' || physical.current?.kind === 'deleting') return;
    setForm((value) => ({ key, value: { ...(value.key === key ? value.value : empty()), ...patch } }));
    clearMessages();
  };
  const startNew = () => {
    if (!current() || busy) return;
    setForm({ key, value: empty(defaultSheet) });
    setEditing(null);
    setMode('new');
    clearMessages();
  };
  const startEdit = (entry: LocalActionEntry) => {
    if (
      !current() ||
      busy ||
      !fresh?.entries.some((value) => value.id === entry.id && value.revision === entry.revision)
    )
      return;
    setForm({ key, value: fromEntry(entry) });
    setEditing({ id: entry.id, revision: entry.revision });
    setMode('edit');
    clearMessages();
  };
  const cancel = () => {
    if (!current() || physical.current?.kind === 'saving' || physical.current?.kind === 'deleting') return;
    setMode('list');
    setEditing(null);
    setForm({ key, value: empty(defaultSheet) });
    clearMessages();
  };
  let checked: LocalActionDraft | null = null,
    reason = '';
  try {
    checked = draft(fields);
  } catch (failure) {
    reason = failure instanceof Error ? failure.message : 'Complete the action definition.';
  }
  async function save(): Promise<void> {
    if (!current() || physical.current || !frame.controls || !fresh || !checked || !editCurrent || mode === 'list')
      return;
    const controls = frame.controls,
      previous = fresh,
      saved = validateLocalActionDraft(checked),
      target = mode === 'edit' ? editing : null,
      job = { kind: 'saving' as const },
      token = ++generation.current;
    physical.current = job;
    setPhase('saving');
    clearMessages();
    try {
      const result = validateLocalActionList(
        await (target
          ? controls.update(target.id, saved, previous.revision)
          : controls.create(saved, previous.revision))
      );
      if (!current() || generation.current !== token) return;
      if (result.revision <= previous.revision) throw Error('The saved library did not advance its registry revision.');
      const match = target
        ? result.entries.find((entry) => entry.id === target.id)
        : result.entries.find((entry) => !previous.entries.some((old) => old.id === entry.id));
      if (
        !match ||
        !equivalent({ name: match.name, descriptor: match.descriptor }, saved) ||
        (target && match.revision <= target.revision)
      )
        throw Error('The saved action reply does not match the submitted definition.');
      setList({ key, value: result });
      setMode('list');
      setEditing(null);
      setForm({ key, value: empty(defaultSheet) });
      setNotice('Action definition saved in the private library.');
    } catch (failure) {
      if (current() && generation.current === token)
        setError(failure instanceof Error ? failure.message : 'The action definition could not be saved.');
    } finally {
      finish(job);
    }
  }
  async function remove(entry: LocalActionEntry): Promise<void> {
    if (
      !current() ||
      physical.current ||
      !frame.controls ||
      !fresh?.entries.some((value) => value.id === entry.id && value.revision === entry.revision)
    )
      return;
    const controls = frame.controls,
      previous = fresh,
      job = { kind: 'deleting' as const },
      token = ++generation.current;
    physical.current = job;
    setPhase('deleting');
    clearMessages();
    try {
      const result = validateLocalActionList(await controls.remove(entry.id, previous.revision));
      if (!current() || generation.current !== token) return;
      if (result.revision <= previous.revision || result.entries.some((value) => value.id === entry.id))
        throw Error('The deleted action reply does not match the submitted entry.');
      setList({ key, value: result });
      setNotice('Action definition deleted from the private library.');
    } catch (failure) {
      if (current() && generation.current === token)
        setError(failure instanceof Error ? failure.message : 'The action definition could not be deleted.');
    } finally {
      finish(job);
    }
  }
  const catalog = localActionCatalog(fields.command);
  const changeCommand = (value: Command) => {
    if (!LOCAL_ACTION_CATALOG.some((entry) => entry.command === value)) return;
    editForm({ ...empty(), name: fields.name, sheet: fields.sheet, command: value });
  };
  const updateKey = (index: number, patch: Partial<Form['keys'][number]>) =>
    editForm({ keys: fields.keys.map((value, i) => (i === index ? { ...value, ...patch } : value)) });
  const moveKey = (index: number, offset: number) => {
    const keys = fields.keys.map((value) => ({ ...value }));
    if (index + offset < 0 || index + offset >= keys.length) return;
    [keys[index], keys[index + offset]] = [keys[index + offset], keys[index]];
    editForm({ keys });
  };
  return (
    <section
      className="sheet-actions"
      aria-label={words('Local Actions')}
      data-formula-draft-zone
      data-action-frame={key}
      data-action-mode={mode}
      data-action-registry-revision={fresh?.revision}
    >
      <p>
        {words(
          'Manage private action definitions. Saving a definition does not run it. Running saved actions is unavailable.'
        )}
      </p>
      {!available ? (
        <p role="status">{words('This host has not supplied a private action library.')}</p>
      ) : !visible ? (
        <p role="status">{words('The private action library is paused while this pane is closed.')}</p>
      ) : (
        <>
          <div>
            <button type="button" disabled={busy} onClick={startNew}>
              {words('New action')}
            </button>
            <button
              type="button"
              disabled={physical.current?.kind === 'saving' || physical.current?.kind === 'deleting'}
              onClick={() => {
                if (current()) {
                  setMode('list');
                  clearMessages();
                }
              }}
            >
              {words('View actions')}
            </button>
            <button type="button" disabled={busy} onClick={() => void load()}>
              {words('Reload library')}
            </button>
          </div>
          {busy && (
            <p role="status">
              {words(
                phase === 'loading'
                  ? 'Loading the private library…'
                  : phase === 'saving'
                    ? 'Saving the action definition…'
                    : 'Deleting the action definition…'
              )}
            </p>
          )}
          {fresh?.warning && <p role="status">{fresh.warning}</p>}
          {mode === 'list' && fresh && (
            <>
              <p>
                {fresh.entries.length} / {LOCAL_ACTION_LIMITS.entries} {words('saved action definitions')}
              </p>
              <ul aria-label={words('Saved actions')}>
                {fresh.entries.map((entry) => (
                  <li key={entry.id} data-action-id={entry.id} data-action-entry-revision={entry.revision}>
                    <strong>{entry.name}</strong>{' '}
                    <span>{words(localActionCatalog(entry.descriptor.command).label)}</span>
                    <button
                      type="button"
                      aria-label={words('Edit action') + ' ' + entry.name}
                      disabled={busy}
                      onClick={() => startEdit(entry)}
                    >
                      {words('Edit')}
                    </button>
                    <button
                      type="button"
                      aria-label={words('Delete action') + ' ' + entry.name}
                      disabled={busy}
                      onClick={() => void remove(entry)}
                    >
                      {words('Delete')}
                    </button>
                  </li>
                ))}
              </ul>
              {!fresh.entries.length && <p>{words('The private library is empty.')}</p>}
            </>
          )}
          {mode !== 'list' && form.key === key && (
            <form
              onSubmit={(event) => {
                event.preventDefault();
                void save();
              }}
            >
              <fieldset disabled={phase === 'saving' || phase === 'deleting'}>
                <legend>{words(mode === 'edit' ? 'Edit action definition' : 'New action definition')}</legend>
                <label>
                  {words('Action name')}
                  <input
                    aria-label={words('Action name')}
                    maxLength={128}
                    value={fields.name}
                    onChange={(event) => editForm({ name: event.target.value })}
                  />
                </label>
                <label>
                  {words('Action command')}
                  <select
                    aria-label={words('Action command')}
                    value={fields.command}
                    onChange={(event) => changeCommand(event.target.value as Command)}
                  >
                    {LOCAL_ACTION_CATALOG.map((entry) => (
                      <option key={entry.command} value={entry.command}>
                        {words(entry.label)}
                      </option>
                    ))}
                  </select>
                </label>
                {fields.command !== 'recalculate' && (
                  <label>
                    {words('Worksheet')}
                    <input
                      aria-label={words('Action worksheet')}
                      maxLength={31}
                      value={fields.sheet}
                      onChange={(event) => editForm({ sheet: event.target.value })}
                    />
                  </label>
                )}
                {fields.command === 'freeze' && (
                  <>
                    <label>
                      {words('Rows to freeze')}
                      <input
                        type="number"
                        min={0}
                        max={100}
                        aria-label={words('Rows to freeze')}
                        value={fields.rows}
                        onChange={(event) => editForm({ rows: event.target.value })}
                      />
                    </label>
                    <label>
                      {words('Columns to freeze')}
                      <input
                        type="number"
                        min={0}
                        max={30}
                        aria-label={words('Columns to freeze')}
                        value={fields.columns}
                        onChange={(event) => editForm({ columns: event.target.value })}
                      />
                    </label>
                    <p>{words('Freeze up to 100 rows and 30 columns. Zero leaves that axis unfrozen.')}</p>
                  </>
                )}
                {(fields.command === 'formatRange' || fields.command === 'sortRows') && (
                  <label>
                    {words('Range')}
                    <input
                      aria-label={words('Action range')}
                      maxLength={64}
                      value={fields.range}
                      onChange={(event) => editForm({ range: event.target.value })}
                    />
                  </label>
                )}
                {fields.command === 'formatRange' && (
                  <>
                    {(['bold', 'italic'] as const).map((field) => (
                      <label key={field}>
                        {words(field === 'bold' ? 'Bold' : 'Italic')}
                        <select
                          aria-label={words('Action ' + field)}
                          value={fields[field]}
                          onChange={(event) => editForm({ [field]: event.target.value as TriState })}
                        >
                          <option value="">{words('Keep unchanged')}</option>
                          <option value="true">{words('On')}</option>
                          <option value="false">{words('Off')}</option>
                        </select>
                      </label>
                    ))}
                    <label>
                      {words('Number format')}
                      <select
                        aria-label={words('Action number format')}
                        value={fields.numberFormat}
                        onChange={(event) => editForm({ numberFormat: event.target.value })}
                      >
                        <option value="">{words('Keep unchanged')}</option>
                        {formats.map((value) => (
                          <option key={value} value={value}>
                            {words(numberLabels[value])}
                          </option>
                        ))}
                      </select>
                    </label>
                    {Object.keys(fields.otherFormat).length > 0 && (
                      <p>
                        {words('Other saved formatting retained')}:{' '}
                        {Object.entries(fields.otherFormat)
                          .map(
                            ([name, value]) =>
                              `${words(formatLabels[name as keyof LocalActionFormat])}: ${String(value)}`
                          )
                          .join(' · ')}
                      </p>
                    )}
                  </>
                )}
                {fields.command === 'sortRows' && (
                  <>
                    <p>{words('Enter data rows only; exclude the header row. Levels are applied in this order.')}</p>
                    <ol aria-label={words('Action sort levels')}>
                      {fields.keys.map((level, index) => (
                        <li key={index}>
                          <label>
                            {words('Column')}
                            <input
                              aria-label={words('Action sort column') + ' ' + (index + 1)}
                              maxLength={3}
                              value={level.column}
                              onChange={(event) => updateKey(index, { column: event.target.value })}
                            />
                          </label>
                          <label>
                            {words('Order')}
                            <select
                              aria-label={words('Action sort order') + ' ' + (index + 1)}
                              value={level.direction}
                              onChange={(event) =>
                                updateKey(index, { direction: event.target.value as 'asc' | 'desc' })
                              }
                            >
                              <option value="asc">{words('Ascending')}</option>
                              <option value="desc">{words('Descending')}</option>
                            </select>
                          </label>
                          <button
                            type="button"
                            aria-label={words('Move action level') + ' ' + (index + 1) + ' ' + words('up')}
                            disabled={index === 0}
                            onClick={() => moveKey(index, -1)}
                          >
                            ↑
                          </button>
                          <button
                            type="button"
                            aria-label={words('Move action level') + ' ' + (index + 1) + ' ' + words('down')}
                            disabled={index === fields.keys.length - 1}
                            onClick={() => moveKey(index, 1)}
                          >
                            ↓
                          </button>
                          <button
                            type="button"
                            aria-label={words('Remove action level') + ' ' + (index + 1)}
                            disabled={fields.keys.length === 1}
                            onClick={() => editForm({ keys: fields.keys.filter((_, i) => i !== index) })}
                          >
                            {words('Remove')}
                          </button>
                        </li>
                      ))}
                    </ol>
                    <button
                      type="button"
                      disabled={fields.keys.length >= 16}
                      onClick={() => editForm({ keys: [...fields.keys, { column: '', direction: 'asc' }] })}
                    >
                      {words('Add action level')}
                    </button>
                  </>
                )}
              </fieldset>
              <p>
                {words('Supported workbook profile')}:{' '}
                {catalog.supportedProfiles.includes('xlsx-writable')
                  ? words('Editable XLSX')
                  : catalog.supportedProfiles.join(', ')}
                . {words('Read scope')}: {words(catalog.readScope)}. {words('Change scope')}:{' '}
                {words(catalog.writeScope)}.
              </p>
              {fields.command === 'formatRange' && (
                <p>
                  {words('Range limit')}: {catalog.budgets.positions.toLocaleString()} {words('positions')}.
                </p>
              )}
              {fields.command === 'sortRows' && (
                <p>
                  {words('Range limit')}: {catalog.budgets.positions.toLocaleString()} {words('positions')} ·{' '}
                  {catalog.budgets.keys} {words('distinct sort columns')}.
                </p>
              )}
              {fields.command === 'recalculate' && (
                <p>{words('Calculation requires supported formulas and available workbook resources.')}</p>
              )}
              {!editCurrent && (
                <p role="status">
                  {words('This saved action changed. Cancel the edit and select its current definition.')}
                </p>
              )}
              {reason && (
                <p role="status" aria-label={words('Action definition issue')}>
                  {words(reason)}
                </p>
              )}
              <button type="submit" disabled={busy || !fresh || !checked || !editCurrent}>
                {words('Save action')}
              </button>
              <button type="button" disabled={phase === 'saving' || phase === 'deleting'} onClick={cancel}>
                {words('Cancel action')}
              </button>
            </form>
          )}
          {error && current() && <p role="alert">{error}</p>}
          {notice && current() && <p role="status">{words(notice)}</p>}
        </>
      )}
    </section>
  );
}
