import { useEffect, useRef, useState } from 'react';
import { pivotFieldDraftReason, type PivotFieldDraft } from './pivotSelection';
import type { WorkbookPivotDefinition } from '../core/contracts';
export function WorkbookPivotContext({
  pivot,
  reason,
  running,
  onRefresh,
  onSource,
  bindingKey,
  sourceSheets = [],
  onUpdate,
  words = (s: string) => s,
}: {
  pivot: WorkbookPivotDefinition;
  reason?: string;
  running: boolean;
  onRefresh: () => Promise<unknown>;
  onSource?: () => void;
  bindingKey?: string;
  sourceSheets?: string[];
  onUpdate?: (draft: PivotFieldDraft) => Promise<unknown>;
  words?: (s: string) => string;
}): JSX.Element {
  const initial = (): PivotFieldDraft => ({
    sourceSheet: pivot.sourceSheet ?? '',
    range: pivot.sourceRange ?? '',
    rowField: pivot.rowField ?? 0,
    valueField: pivot.valueField ?? 1,
    aggregate: pivot.aggregate ?? 'sum',
  });
  const [draft, setDraft] = useState(initial),
    live = useRef({ key: bindingKey, draft });
  live.current = { key: bindingKey, draft };
  useEffect(() => setDraft(initial()), [bindingKey]);
  const patch = (value: Partial<PivotFieldDraft>) => setDraft((previous) => ({ ...previous, ...value }));
  const draftReason = pivotFieldDraftReason(draft),
    unchanged = JSON.stringify(draft) === JSON.stringify(initial());
  const field = (index: number | undefined) =>
    index === undefined ? words('Unavailable') : (pivot.fieldNames?.[index] ?? words('Field') + ' ' + (index + 1));
  return (
    <section
      aria-label={words('PivotTable Fields')}
      data-pivot-name={pivot.name}
      data-pivot-revision={pivot.revision ?? ''}
    >
      <h3>{pivot.name}</h3>
      <dl>
        <dt>{words('Output')}</dt>
        <dd>
          {pivot.sheet}!{pivot.outputRange}
        </dd>
        <dt>{words('Source')}</dt>
        <dd>
          {pivot.sourceSheet && pivot.sourceRange
            ? pivot.sourceSheet + '!' + pivot.sourceRange
            : words('Unavailable for this protected pivot')}
        </dd>
        <dt>{words('Rows')}</dt>
        <dd>{field(pivot.rowField)}</dd>
        <dt>{words('Values')}</dt>
        <dd>
          {pivot.aggregate ? pivot.aggregate.toUpperCase() + ' · ' : ''}
          {field(pivot.valueField)}
        </dd>
      </dl>
      {!!pivot.fieldNames?.length && (
        <ul aria-label={words('Pivot source fields')}>
          {pivot.fieldNames.map((name, index) => (
            <li key={index}>{name}</li>
          ))}
        </ul>
      )}
      {reason && <p role="status">{words(reason)}</p>}
      <button
        type="button"
        disabled={!!reason || running}
        title={reason ? words(reason) : undefined}
        onClick={() => void onRefresh()}
      >
        {words(running ? 'Refreshing pivot…' : 'Refresh selected pivot')}
      </button>
      <button type="button" disabled={!onSource || running} onClick={onSource}>
        {words('Go to pivot source')}
      </button>
      {onUpdate && (
        <fieldset disabled={!!reason || running}>
          <legend>{words('Edit pivot fields')}</legend>
          <label>
            {words('Source worksheet')}
            <select
              aria-label={words('Pivot source worksheet')}
              value={draft.sourceSheet}
              onChange={(e) => patch({ sourceSheet: e.target.value })}
            >
              {sourceSheets
                .filter((name) => name !== pivot.sheet)
                .map((name) => (
                  <option key={name}>{name}</option>
                ))}
            </select>
          </label>
          <label>
            {words('Source rectangle')}
            <input
              aria-label={words('Pivot source range')}
              value={draft.range}
              onChange={(e) => patch({ range: e.target.value })}
            />
          </label>
          <label>
            {words('Row field number')}
            <input
              type="number"
              min="1"
              max="8"
              aria-label={words('Pivot row field number')}
              value={draft.rowField + 1}
              onChange={(e) => patch({ rowField: Number(e.target.value) - 1 })}
            />
          </label>
          <label>
            {words('Value field number')}
            <input
              type="number"
              min="1"
              max="8"
              aria-label={words('Pivot value field number')}
              value={draft.valueField + 1}
              onChange={(e) => patch({ valueField: Number(e.target.value) - 1 })}
            />
          </label>
          <label>
            {words('Aggregate')}
            <select
              aria-label={words('Pivot value aggregate')}
              value={draft.aggregate}
              onChange={(e) => patch({ aggregate: e.target.value as 'sum' | 'count' })}
            >
              <option value="sum">SUM</option>
              <option value="count">{words('Numeric COUNT')}</option>
            </select>
          </label>
          <small>
            {words(
              'Field numbers follow the headers in the chosen source rectangle. Apply rebuilds the pivot from current qualified inputs.'
            )}
          </small>
          {draftReason && <p role="status">{words(draftReason)}</p>}
          <button
            type="button"
            disabled={!!draftReason || unchanged}
            onClick={() => {
              if (live.current.key === bindingKey && live.current.draft === draft)
                void onUpdate(structuredClone(draft));
            }}
          >
            {words(running ? 'Updating pivot…' : 'Apply pivot fields')}
          </button>
        </fieldset>
      )}
      <p>
        {words(
          'Refresh updates the displayed pivot from its supported source. Selection and field inspection leave the workbook unchanged.'
        )}
      </p>
    </section>
  );
}
