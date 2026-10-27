import { useEffect, useRef, useState } from 'react';
import type { WorkbookCommand, WorkbookView } from '../core/contracts';
export interface WorkbookReferencesProps {
  sheet: string;
  selection: string;
  revision?: string;
  tables: NonNullable<WorkbookView['tables']>;
  pivots: NonNullable<WorkbookView['pivots']>;
  readOnly: boolean;
  busy: boolean;
  active: boolean;
  onBatch: (commands: WorkbookCommand[]) => void | Promise<unknown>;
}
/** Commands carry the definition displayed by the user; the owner rechecks it. */
export function WorkbookReferences(p: WorkbookReferencesProps): JSX.Element {
  const [insertAt, setInsertAt] = useState(0),
    [addedLabel, setAddedLabel] = useState(''),
    [column, setColumn] = useState(''),
    [columnLabel, setColumnLabel] = useState(''),
    [selected, setSelected] = useState(''),
    [name, setName] = useState(''),
    [range, setRange] = useState(p.selection),
    [pivotName, setPivotName] = useState('SummaryPivot'),
    [target, setTarget] = useState('Summary'),
    [row, setRow] = useState(1),
    [value, setValue] = useState(2),
    [aggregate, setAggregate] = useState<'sum' | 'count'>('sum'),
    [error, setError] = useState(''),
    [running, setRunning] = useState(false);
  const binding = JSON.stringify([p.revision, p.sheet, p.selection, p.active, p.readOnly]),
    guard = JSON.stringify([binding, p.busy]);
  const live = useRef({ binding, guard, epoch: 0, active: p.active, readOnly: p.readOnly, busy: p.busy }),
    writing = useRef(false),
    mounted = useRef(true);
  if (live.current.guard !== guard)
    live.current = {
      binding,
      guard,
      epoch: live.current.epoch + 1,
      active: p.active,
      readOnly: p.readOnly,
      busy: p.busy,
    };
  const generation = live.current.epoch;
  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
    };
  }, []);
  const table = p.tables.find((t) => t.name === selected && t.sheet === p.sheet),
    disabled = p.readOnly || p.busy || running || !p.active;
  useEffect(() => {
    setColumn('');
    setColumnLabel('');
    setInsertAt(table?.columns.length ?? 0);
    setAddedLabel('');
    setRange(table?.ref ?? p.selection);
  }, [selected, p.sheet, p.revision]);
  useEffect(() => {
    setSelected('');
    setName('');
    setRange(p.selection);
    setError('');
  }, [p.sheet]);
  async function send(command: WorkbookCommand): Promise<void> {
    const current = live.current;
    if (
      !mounted.current ||
      disabled ||
      writing.current ||
      current.epoch !== generation ||
      current.guard !== guard ||
      !current.active ||
      current.readOnly ||
      current.busy
    )
      return;
    const scope = binding;
    writing.current = true;
    setRunning(true);
    setError('');
    try {
      await p.onBatch([structuredClone(command)]);
    } catch (e) {
      if (mounted.current && live.current.binding === scope) setError(e instanceof Error ? e.message : String(e));
    } finally {
      writing.current = false;
      if (mounted.current) setRunning(false);
    }
  }
  return (
    <section className="sheet-name-manager" aria-label="Tables and pivots">
      <p>
        Ordinary tables support table and column rename, plus row resize with fixed header origin. Add/delete columns at
        any position within the table; adding requires blank space immediately to its right. Pivot creation uses one
        text row field and SUM or numeric COUNT of integer values, on a new sheet. Refresh is explicit.
      </p>
      <fieldset disabled={disabled}>
        <legend>Table editing</legend>
        <label>
          Table{' '}
          <select
            aria-label="Table to edit"
            value={selected}
            onChange={(e) => {
              setSelected(e.target.value);
              const t = p.tables.find((t) => t.name === e.target.value && t.sheet === p.sheet);
              setName(t?.name ?? '');
              setRange(t?.ref ?? p.selection);
            }}
          >
            <option value="">Select a table on this sheet</option>
            {p.tables
              .filter((t) => t.sheet === p.sheet)
              .map((t) => (
                <option key={t.name}>{t.name}</option>
              ))}
          </select>
        </label>
        <label>
          New table name <input aria-label="New table name" value={name} onChange={(e) => setName(e.target.value)} />
        </label>
        <button
          disabled={!table || name === table.name}
          onClick={() =>
            table &&
            void send({
              operation: 'renameTable',
              sheet: p.sheet,
              tableName: table.name,
              expectedTableRef: table.ref,
              newName: name.trim(),
            })
          }
        >
          Rename table
        </button>
        <label>
          Table range <input aria-label="Table range" value={range} onChange={(e) => setRange(e.target.value)} />
        </label>
        <button onClick={() => setRange(p.selection)}>Use selected table range</button>
        <button
          disabled={!table || range === table.ref}
          onClick={() =>
            table &&
            void send({
              operation: 'resizeTable',
              sheet: p.sheet,
              tableName: table.name,
              expectedTableRef: table.ref,
              range,
            })
          }
        >
          Resize table rows
        </button>
        <label>
          Column{' '}
          <select
            aria-label="Table column to rename"
            value={column}
            onChange={(e) => {
              setColumn(e.target.value);
              setColumnLabel(e.target.value);
            }}
          >
            <option value="">Select a column</option>
            {table?.columns.map((c) => (
              <option key={c.name}>{c.name}</option>
            ))}
          </select>
        </label>
        <label>
          New column label{' '}
          <input
            aria-label="New table column label"
            value={columnLabel}
            onChange={(e) => setColumnLabel(e.target.value)}
          />
        </label>
        <button
          disabled={!table || !column || !columnLabel.trim() || columnLabel === column}
          onClick={() =>
            table &&
            void send({
              operation: 'renameTableColumn',
              sheet: p.sheet,
              tableName: table.name,
              expectedTableRef: table.ref,
              columnName: column,
              newName: columnLabel,
            })
          }
        >
          Rename table column
        </button>
        <label>
          Insert position{' '}
          <select
            aria-label="Insert table column position"
            value={insertAt}
            onChange={(e) => setInsertAt(Number(e.target.value))}
          >
            {table?.columns.map((c, i) => (
              <option key={c.name} value={i}>
                Before {c.name}
              </option>
            ))}
            <option value={table?.columns.length ?? 0}>At end</option>
          </select>
        </label>
        <label>
          Added column label{' '}
          <input
            aria-label="Added table column label"
            value={addedLabel}
            onChange={(e) => setAddedLabel(e.target.value)}
          />
        </label>
        <button
          disabled={!table || !addedLabel.trim()}
          onClick={() =>
            table &&
            void send({
              operation: 'insertTableColumn',
              sheet: p.sheet,
              tableName: table.name,
              expectedTableRef: table.ref,
              index: insertAt,
              newName: addedLabel,
            })
          }
        >
          Add table column
        </button>
        <button
          disabled={!table || !column || table.columns.length <= 1}
          onClick={() =>
            table &&
            void send({
              operation: 'deleteTableColumn',
              sheet: p.sheet,
              tableName: table.name,
              expectedTableRef: table.ref,
              columnName: column,
            })
          }
        >
          Delete table column
        </button>
        <small>
          Delete removes the selected column's cells within this table. Undo restores them. Keep at least one column.
          Cells outside the table stay in place; ranges crossing the table movement boundary are protected.
        </small>
        <small>
          Column rename updates the header and ordinary formula references together. Labels with brackets or selector
          punctuation and unsupported dependency graphs are protected.
        </small>
      </fieldset>
      <fieldset disabled={disabled || p.pivots.length > 0}>
        <legend>Create pivot</legend>
        <p>
          Source: {p.sheet}!{p.selection}. The first row supplies unique text headers. Maximum 8,000 positions, 8 fields
          and 512 row groups.
        </p>
        <label>
          Pivot name <input aria-label="Pivot name" value={pivotName} onChange={(e) => setPivotName(e.target.value)} />
        </label>
        <label>
          New pivot sheet{' '}
          <input aria-label="New pivot sheet" value={target} onChange={(e) => setTarget(e.target.value)} />
        </label>
        <label>
          Row field number{' '}
          <input
            aria-label="Pivot row field number"
            type="number"
            min="1"
            max="8"
            value={row}
            onChange={(e) => setRow(Number(e.target.value))}
          />
        </label>
        <label>
          Value field number{' '}
          <input
            aria-label="Pivot value field number"
            type="number"
            min="1"
            max="8"
            value={value}
            onChange={(e) => setValue(Number(e.target.value))}
          />
        </label>
        <label>
          Aggregate{' '}
          <select
            aria-label="Pivot aggregate"
            value={aggregate}
            onChange={(e) => setAggregate(e.target.value as 'sum' | 'count')}
          >
            <option value="sum">SUM</option>
            <option value="count">COUNT of numeric values</option>
          </select>
        </label>
        <button
          onClick={() =>
            void send({
              operation: 'createPivot',
              sheet: p.sheet,
              name: pivotName.trim(),
              range: p.selection,
              rowField: row - 1,
              valueField: value - 1,
              aggregate,
              pivotSheet: target.trim(),
              expectedPivotRevision: null,
            })
          }
        >
          Create pivot
        </button>
      </fieldset>
      <div aria-label="Pivot inventory">
        {p.pivots.map((pivot, i) => (
          <div key={pivot.name + '\0' + i}>
            <strong>{pivot.name}</strong>{' '}
            <span>
              {pivot.sheet ?? 'Unknown sheet'} · {pivot.sourceSheet}!{pivot.sourceRange}
            </span>
            {pivot.canEdit && pivot.revision && pivot.sheet ? (
              <button
                disabled={disabled}
                onClick={() =>
                  void send({
                    operation: 'refreshPivot',
                    sheet: pivot.sheet!,
                    name: pivot.name,
                    expectedPivotRevision: pivot.revision!,
                  })
                }
              >
                Refresh pivot {pivot.name}
              </button>
            ) : (
              <p>{pivot.reason ?? 'Imported pivot is preserved and protected.'}</p>
            )}
          </div>
        ))}
      </div>
      {running && <p role="status">Updating workbook definitions…</p>}
      {error && <p role="alert">{error}</p>}
    </section>
  );
}
