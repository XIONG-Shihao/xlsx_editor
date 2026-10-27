import assert from 'node:assert/strict';
import { refusal } from '../core/refusal-codes.mjs';

const own = (value, key) => Object.hasOwn(value, key);
const text = (value) => typeof value === 'string' && value.length > 0 && value.length <= 4096 && !value.includes('\0');
const exact = (value, keys) =>
  value &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  Object.keys(value).length === keys.length &&
  keys.every((k) => own(value, k));
export function validateCalculationAction(value) {
  if (value?.type === 'calculateAllOpen' && exact(value, ['type'])) return { type: value.type };
  if (value?.type === 'calculateSheet' && exact(value, ['type', 'sheet']) && text(value.sheet))
    return { type: value.type, sheet: value.sheet };
  if (
    value?.type === 'setOptions' &&
    ['auto', 'manual'].includes(value.mode) &&
    exact(value, own(value, 'recalculateBeforeSave') ? ['type', 'mode', 'recalculateBeforeSave'] : ['type', 'mode']) &&
    (!own(value, 'recalculateBeforeSave') || typeof value.recalculateBeforeSave === 'boolean')
  )
    return { ...value };
  throw refusal('formula-unsupported', 'Unsupported calculation action or fields.');
}
function binding(value) {
  assert(
    value && ['documentId', 'ownerId', 'state', 'revision'].every((k) => text(value[k])),
    'Calculation requires an accepted workbook binding.'
  );
  return Object.fromEntries(['documentId', 'ownerId', 'state', 'revision'].map((k) => [k, value[k]]));
}
/** The host captures private participants; no renderer-supplied roster or path
 * enters this operation. Each workbook has an independent durable commit. */
export async function runCalculationGroup({
  requestId,
  initiator,
  action: incoming,
  targets,
  current,
  authorize,
  apply,
  signal,
  progress,
  onAccepted,
  maxParticipants = 32,
}) {
  assert(
    typeof requestId === 'string' && /^[a-f0-9]{8}-[a-f0-9-]{27}$/i.test(requestId),
    'Invalid calculation job identity.'
  );
  const action = validateCalculationAction(incoming),
    start = binding(initiator);
  assert(
    Array.isArray(targets) && targets.length > 0 && targets.length <= maxParticipants,
    'Too many or no granted open workbooks.'
  );
  assert(Number.isInteger(maxParticipants) && maxParticipants >= 1 && maxParticipants <= 128);
  const seen = new Set(),
    captured = targets.map((target) => {
      const accepted = binding(target);
      assert(
        text(target.name) &&
          ['xlsx', 'delimited', 'passive'].includes(target.profile) &&
          typeof target.readOnly === 'boolean'
      );
      assert(!seen.has(accepted.ownerId), 'Duplicate calculation owner.');
      seen.add(accepted.ownerId);
      return { target, before: accepted };
    });
  const first = captured.find((c) => c.before.ownerId === start.ownerId && c.before.documentId === start.documentId);
  assert(
    first &&
      first.before.state === start.state &&
      first.before.revision === start.revision &&
      current(first.target, start.state),
    'The initiating calculation owner changed.'
  );
  const selected = action.type === 'calculateSheet' ? [first] : [first, ...captured.filter((c) => c !== first)];
  const results = [];
  let cancelled = false;
  const emit = (phase, completed, currentName, cancellable) => {
    try {
      progress?.({
        requestId,
        phase,
        completed,
        total: selected.length,
        ...(currentName ? { currentName } : {}),
        cancellable,
      });
    } catch {
      /* Feedback cannot reverse a durable owner outcome. */
    }
  };
  for (const { target, before } of selected) {
    const row = {
      documentId: before.documentId,
      ownerId: before.ownerId,
      name: target.name,
      beforeState: before.state,
    };
    if (cancelled || signal?.aborted) {
      cancelled = true;
      results.push({ ...row, status: 'notRun', reason: 'Calculation cancelled before this workbook started.' });
      continue;
    }
    let reason =
      target.profile !== 'xlsx'
        ? 'This format preserves stored results.'
        : target.readOnly
          ? (target.reason ?? 'This workbook is read-only.')
          : action.type === 'calculateAllOpen' && target.canRecalculate === false
            ? (target.reason ?? 'Calculation is protected for this workbook.')
            : action.type === 'setOptions' && target.canChangeOptions !== true
              ? (target.reason ?? 'Calculation settings are unavailable for this workbook.')
              : action.type === 'calculateSheet' && target.canCalculateSheet !== true
                ? (target.reason ?? 'Scoped calculation is unavailable for this workbook.')
                : undefined;
    if (reason) {
      results.push({ ...row, status: 'skipped', reason });
      continue;
    }
    if (!current(target, before.state)) {
      results.push({ ...row, status: 'refused', reason: 'The workbook owner or accepted state changed.' });
      continue;
    }
    emit('preparing', results.length, target.name, true);
    try {
      await authorize(target);
    } catch (error) {
      results.push({ ...row, status: 'refused', reason: String(error.message ?? error) });
      continue;
    }
    if (signal?.aborted) {
      cancelled = true;
      results.push({ ...row, status: 'notRun', reason: 'Calculation cancelled before acceptance.' });
      continue;
    }
    if (!current(target, before.state)) {
      results.push({ ...row, status: 'refused', reason: 'The workbook changed during authorization.' });
      continue;
    }
    const command =
      action.type === 'calculateAllOpen'
        ? { type: 'recalculate' }
        : action.type === 'calculateSheet'
          ? { type: 'calculateSheet', sheet: action.sheet }
          : {
              type: 'setCalculationOptions',
              mode: action.mode,
              ...(own(action, 'recalculateBeforeSave') ? { recalculateBeforeSave: action.recalculateBeforeSave } : {}),
            };
    emit(action.type === 'setOptions' ? 'updating settings' : 'calculating', results.length, target.name, true);
    try {
      const reply = binding(await apply(target, [command]));
      if (reply.documentId !== before.documentId || reply.ownerId !== before.ownerId)
        throw refusal('formula-stale', 'Calculation receipt belongs to another workbook owner.');
      // A durable acceptance remains accepted if Cancel arrives just afterward.
      const accepted = { ...row, status: 'accepted', afterState: reply.state };
      results.push(accepted);
      if (current(target, reply.state)) {
        try {
          onAccepted?.(target, reply);
        } catch (error) {
          accepted.reason = 'Accepted; host feedback failed: ' + String(error.message ?? error);
        }
      } else accepted.reason = 'Accepted in the original workbook; its view has since retired.';
    } catch (error) {
      if (['CALCULATION_CANCEL_GRANTED', 'OWNER_COMMAND_CANCELLED'].includes(error?.code)) {
        cancelled = true;
        results.push({
          ...row,
          status: 'cancelled',
          reason: 'The current workbook calculation was cancelled; earlier accepted workbooks were kept.',
        });
      } else results.push({ ...row, status: 'failed', reason: String(error.message ?? error) });
    }
  }
  const outcome =
    cancelled || results.some((r) => ['cancelled', 'notRun'].includes(r.status))
      ? 'cancelled'
      : results.some((r) => ['failed', 'refused'].includes(r.status))
        ? 'partial'
        : 'completed';
  emit(outcome, results.length, undefined, false);
  return {
    schema: 1,
    requestId,
    initiator: start,
    action,
    consistency: 'independent-workbook-commits',
    outcome,
    results,
  };
}
