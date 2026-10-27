import { useId, useSyncExternalStore } from 'react';
import {
  calculationDisabledReason,
  calculationPendingObjectCount,
  calculationFrameKey,
  CalculationJobRetired,
  type CalculationFrame,
  type CalculationJobController,
  type CalculationJobSnapshot,
} from './calculationJobs';
export interface WorkbookCalculationProps {
  controller: CalculationJobController;
  frame: CalculationFrame;
  words?: (value: string) => string;
}
const phaseText = (snapshot: CalculationJobSnapshot): string => {
  if (snapshot.phase === 'stopping') return 'Cancellation requested. Waiting for the current workbook result…';
  if (snapshot.phase === 'finishing') return 'Finishing calculation. Accepted changes will be kept.';
  if (snapshot.phase === 'running')
    return snapshot.action?.type === 'setOptions'
      ? 'Updating calculation options…'
      : snapshot.action?.type === 'calculateSheet'
        ? `Calculating worksheet ${snapshot.action.sheet}…`
        : 'Calculating open workbooks…';
  if (snapshot.phase === 'failed') return 'Calculation was not completed.';
  if (snapshot.phase === 'cancelled') return 'Calculation cancelled. Earlier accepted changes were kept.';
  if (snapshot.phase === 'partial') return 'Some workbooks could not finish. Accepted changes were kept.';
  if (snapshot.phase === 'completed')
    return snapshot.receipt?.results.every((row) => row.status === 'skipped')
      ? 'No applicable workbook changes.'
      : snapshot.action?.type === 'setOptions'
        ? 'Calculation options updated.'
        : 'Calculation finished.';
  return '';
};
export function WorkbookCalculationStatus({
  controller,
  words = (text: string) => text,
}: {
  controller: CalculationJobController;
  words?: (value: string) => string;
}): JSX.Element | null {
  useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const job = controller.visible();
  if (job.phase === 'idle') return null;
  return (
    <div
      className="sheet-view-notice"
      role="status"
      aria-label={words('Calculation status')}
      data-calculation-phase={job.phase}
      data-calculation-request={job.requestId ?? undefined}
    >
      {words(phaseText(job))}{' '}
      {job.progress && (
        <span aria-label={words('Calculation progress')}>
          {job.progress.completed}/{job.progress.total}
          {job.progress.currentName ? ` · ${job.progress.currentName}` : ''}
        </span>
      )}
      {job.physicalBusy && (
        <button
          type="button"
          aria-label={words('Cancel calculation')}
          disabled={!job.canCancel}
          onClick={() => void controller.cancel(job.requestId!).catch(() => {})}
        >
          {words(
            job.phase === 'finishing' ? 'Finishing' : job.phase === 'stopping' ? 'Stopping' : 'Cancel calculation'
          )}
        </button>
      )}
    </div>
  );
}
export function WorkbookCalculation({
  controller,
  frame,
  words = (text: string) => text,
}: WorkbookCalculationProps): JSX.Element {
  useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const groupId = useId(),
    job = controller.visible(),
    policy = frame.controls.policy,
    key = calculationFrameKey(frame),
    app = frame.controls.applicationMode,
    pendingObjects = calculationPendingObjectCount(policy),
    selectedMode = frame.controls.optionsScope === 'application' && app ? app.mode : policy.mode;
  const action = (mode: 'auto' | 'manual', recalculateBeforeSave?: boolean) => ({
    type: 'setOptions' as const,
    mode,
    ...(recalculateBeforeSave === undefined ? {} : { recalculateBeforeSave }),
  });
  const reason = calculationDisabledReason(frame, action(selectedMode)),
    disabled = !!reason || job.physicalBusy;
  const submit = (mode: 'auto' | 'manual', save?: boolean) =>
    void controller.run(action(mode, save), key).catch((error) => {
      if (!(error instanceof CalculationJobRetired)) console.warn('xlsx.calculation-options', String(error));
    });
  return (
    <section
      aria-label={words('Calculation Options')}
      data-calculation-policy-state={frame.scope.state}
      className="sheet-calculation-options"
    >
      <p>
        {words('Current workbook:')} <strong>{words(policy.mode === 'manual' ? 'Manual' : 'Automatic')}</strong>
      </p>
      <p role="status" aria-label={words('Pending formulas')} data-pending-formulas={policy.pendingFormulaCells}>
        {policy.pendingFormulaCells} {words('formulas pending calculation')}.{' '}
        {policy.pendingFormulaCells > 0 && words('Stored results are not current calculation results.')}
      </p>
      {!!pendingObjects && (
        <p role="status" aria-label={words('Pending object caches')} data-pending-cache-objects={pendingObjects}>
          {pendingObjects} {words('chart/pivot caches pending calculation. Stored output is not current.')}
        </p>
      )}
      {frame.controls.optionsScope === 'application' && (
        <p>
          {words(
            'These options apply to open editable workbooks and future workbooks. Each workbook commits separately.'
          )}
        </p>
      )}
      {app && (
        <p aria-label={words('Application calculation mode')}>
          {words('Application default:')} {words(app.mode === 'manual' ? 'Manual' : 'Automatic')}
          {app.mode !== policy.mode && `. ${words('This workbook currently has a different calculation mode.')}`}
        </p>
      )}
      <fieldset disabled={disabled}>
        <legend>{words('Calculation mode')}</legend>
        <label>
          <input
            type="radio"
            name={groupId}
            aria-label={words('Automatic')}
            checked={selectedMode === 'auto'}
            onChange={() => submit('auto')}
          />
          {words('Automatic')}
        </label>
        <label>
          <input
            type="radio"
            name={groupId}
            aria-label={words('Manual')}
            checked={selectedMode === 'manual'}
            onChange={() => submit('manual')}
          />
          {words('Manual')}
        </label>
      </fieldset>
      <label>
        <input
          type="checkbox"
          aria-label={words('Recalculate before saving')}
          checked={policy.recalculateBeforeSave}
          disabled={disabled}
          onChange={(event) => submit(selectedMode, event.target.checked)}
        />
        {words('Recalculate before saving')}
      </label>
      {frame.controls.optionsScope === 'application' && app && app.mode !== policy.mode && (
        <small>{words('Changing these options uses the application default mode shown above.')}</small>
      )}
      {reason && <p role="status">{words(reason)}</p>}
      {job.error && <p role="alert">{words(job.error)}</p>}
      {job.receipt && (
        <section aria-label={words('Calculation results')} data-calculation-outcome={job.receipt.outcome}>
          <h3>{words('Last calculation results')}</h3>
          <p>
            {job.receipt.results.filter((row) => row.status === 'accepted').length} {words('workbooks accepted')};{' '}
            {job.receipt.results.filter((row) => row.status === 'skipped').length} {words('skipped')}.
          </p>
          <p>{words('Each accepted workbook change has its own Undo step.')}</p>
          <ul>
            {job.receipt.results.map((row) => (
              <li
                key={`${row.documentId}:${row.ownerId}`}
                data-calculation-document={row.documentId}
                data-calculation-owner={row.ownerId}
                data-calculation-result={row.status}
              >
                <strong>{row.name}</strong>: {words(row.status === 'notRun' ? 'not run' : row.status)}
                {row.reason ? ` — ${words(row.reason)}` : ''}
              </li>
            ))}
          </ul>
        </section>
      )}
    </section>
  );
}
