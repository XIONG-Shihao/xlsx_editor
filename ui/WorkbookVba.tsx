import { useEffect, useRef, useState } from 'react';
import type {
  PassiveVbaPage,
  PassiveVbaSource,
  PureVbaFunctionPage,
  PureVbaEvaluation,
} from '../adapters/passive-workbook.mjs';
export interface WorkbookVbaProps {
  scope: string;
  busy: boolean;
  request: (request: {
    operation: 'vbaModules' | 'vbaModule' | 'vbaFunctions' | 'evaluateVba';
    moduleId?: string;
    offset?: number;
    limit?: number;
    functionName?: string;
    args?: number[];
  }) => Promise<any>;
  /** Only for old-format workbooks, which cannot be edited; an editable macro workbook is saved as it is. */
  exportValues?: (
    format: 'xls' | 'xlsb' | 'xlsx'
  ) => Promise<{ copied?: boolean; cancelled?: boolean; omissions?: string[]; revision?: string }>;
}
export function WorkbookVba(p: WorkbookVbaProps): JSX.Element {
  const exportValues = p.exportValues;
  const [open, setOpen] = useState(false),
    [modules, setModules] = useState<PassiveVbaPage>(),
    [source, setSource] = useState<PassiveVbaSource>(),
    [functions, setFunctions] = useState<PureVbaFunctionPage>(),
    [selected, setSelected] = useState(''),
    [fn, setFn] = useState(''),
    [args, setArgs] = useState(''),
    [result, setResult] = useState<PureVbaEvaluation>(),
    [error, setError] = useState(''),
    [running, setRunning] = useState(false),
    [format, setFormat] = useState<'xls' | 'xlsb' | 'xlsx'>('xlsx'),
    [exported, setExported] = useState('');
  const live = useRef(p.scope);
  live.current = p.scope;
  useEffect(() => {
    setRunning(false);
    setOpen(false);
    setArgs('');
    setModules(undefined);
    setSource(undefined);
    setFunctions(undefined);
    setResult(undefined);
    setSelected('');
    setFn('');
    setError('');
    setExported('');
  }, [p.scope]);
  async function work(task: () => Promise<void>): Promise<void> {
    if (p.busy || running) return;
    const scope = p.scope;
    setRunning(true);
    setError('');
    try {
      await task();
    } catch (e) {
      if (live.current === scope) setError(e instanceof Error ? e.message : String(e));
    } finally {
      if (live.current === scope) setRunning(false);
    }
  }
  async function list(offset = 0): Promise<void> {
    const scope = p.scope,
      page: PassiveVbaPage = await p.request({ operation: 'vbaModules', offset, limit: 20 });
    if (live.current === scope) setModules(page);
  }
  async function load(id: string, offset = 0): Promise<void> {
    const scope = p.scope;
    const page: PassiveVbaSource = await p.request({ operation: 'vbaModule', moduleId: id, offset, limit: 32768 });
    const f: PureVbaFunctionPage = await p.request({ operation: 'vbaFunctions', moduleId: id });
    if (live.current === scope) {
      setSource(page);
      setFunctions(f);
      setSelected(id);
      setFn(f.functions[0]?.name ?? '');
      setResult(undefined);
    }
  }
  const disabled = p.busy || running,
    chosen = functions?.functions.find((f) => f.name === fn);
  return (
    <section className="sheet-name-manager" aria-label="VBA and binary workbook actions">
      <p>
        Formula results are stored values from the original file. Save Copy preserves the original bytes. A new
        values-only file omits formulas, macros, formatting and workbook objects.
      </p>
      <button
        disabled={disabled}
        onClick={() => {
          setOpen((v) => !v);
          if (!open) void work(() => list());
        }}
      >
        Inspect VBA source
      </button>
      {open && (
        <>
          <p>
            VBA events and macros do not run. Pure numeric evaluation is a separate, explicit action on an admitted
            function; it does not change cells.
          </p>
          {modules?.protectedReason && <p>{modules.protectedReason}</p>}
          <select
            aria-label="VBA module"
            disabled={disabled}
            value={selected}
            onChange={(e) => void work(() => load(e.target.value))}
          >
            <option value="">Select a module</option>
            {modules?.modules.map((m) => (
              <option key={m.id} value={m.id}>
                {m.name} · {m.kind}
              </option>
            ))}
          </select>
          {modules?.nextOffset !== null && modules?.nextOffset !== undefined && (
            <button disabled={disabled} onClick={() => void work(() => list(modules.nextOffset!))}>
              More VBA modules
            </button>
          )}
          {source && (
            <>
              <p>
                {source.name} · source SHA256 {source.sha256} · {source.offset}–{source.offset + source.code.length} /{' '}
                {source.total}
              </p>
              <pre style={{ maxHeight: '30vh', overflow: 'auto', whiteSpace: 'pre-wrap' }}>{source.code}</pre>
              {source.nextOffset !== null && (
                <button disabled={disabled} onClick={() => void work(() => load(selected, source.nextOffset!))}>
                  More VBA source
                </button>
              )}
            </>
          )}
          {functions?.protectedReason && <p>{functions.protectedReason}</p>}
          {!!functions?.functions.length && (
            <>
              <select
                aria-label="Pure numeric function"
                value={fn}
                disabled={disabled}
                onChange={(e) => {
                  setFn(e.target.value);
                  setResult(undefined);
                }}
              >
                {functions.functions.map((f) => (
                  <option key={f.name} value={f.name}>
                    {f.name}({f.parameters.join(', ')})
                  </option>
                ))}
              </select>
              <label>
                Numeric arguments{' '}
                <input
                  aria-label="Pure numeric arguments"
                  value={args}
                  disabled={disabled}
                  placeholder="2, 3"
                  onChange={(e) => setArgs(e.target.value)}
                />
              </label>
              <button
                disabled={disabled || !chosen}
                onClick={() =>
                  void work(async () => {
                    const scope = p.scope,
                      parsed = args.trim() ? args.split(',').map((v) => Number(v.trim())) : [];
                    if (
                      (args.trim() && args.split(',').some((v) => !v.trim())) ||
                      parsed.length !== chosen?.parameters.length ||
                      parsed.some((n) => !Number.isFinite(n))
                    )
                      throw Error('Supply one finite numeric argument for each parameter.');
                    const evaluated: PureVbaEvaluation = await p.request({
                      operation: 'evaluateVba',
                      moduleId: selected,
                      functionName: fn,
                      args: parsed,
                    });
                    if (live.current === scope) setResult(evaluated);
                  })
                }
              >
                Evaluate pure numeric function
              </button>
              {result && (
                <p role="status">
                  {result.functionName} = {result.value} · {result.steps} interpreter steps
                </p>
              )}
            </>
          )}
        </>
      )}
      {exportValues && (
        <>
          <label>
            New values-only format{' '}
            <select
              aria-label="Values-only output format"
              disabled={disabled}
              value={format}
              onChange={(e) => setFormat(e.target.value as typeof format)}
            >
              <option value="xlsx">XLSX</option>
              <option value="xls">XLS</option>
              <option value="xlsb">XLSB</option>
            </select>
          </label>
          <button
            disabled={disabled}
            onClick={() =>
              void work(async () => {
                const scope = p.scope,
                  r = await exportValues(format);
                if (live.current === scope)
                  setExported(
                    r.cancelled
                      ? 'Export cancelled.'
                      : r.copied
                        ? 'New values-only file saved. ' + (r.omissions ?? []).join(', ')
                        : 'No file was saved.'
                  );
              })
            }
          >
            Export new values-only file…
          </button>
        </>
      )}
      {exported && <p role="status">{exported}</p>}
      {error && <p role="alert">{error}</p>}
    </section>
  );
}
