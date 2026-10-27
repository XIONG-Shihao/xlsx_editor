/** Worker-side cancellation fence. A grant locks out every later critical
 * boundary until the host has stopped/reaped this owner. UI phases are advisory. */
import { refusal } from '../core/refusal-codes.mjs';

export function operationGate(requestId, emit = () => {}) {
  let locked = false,
    granted = false,
    ended = false,
    safe = true,
    last = '',
    events = 0;
  const stopped = () => {
    const error = refusal('operation-cancelled', 'The spreadsheet operation was stopped before its next commit.');
    error.code = 'OPERATION_CANCELLED';
    return error;
  };
  const gate = {
    progress(event) {
      safe = event.cancellable !== false && !locked && !granted && !ended;
      const key = event.phase + ':' + safe + ':' + (event.recoveryId ?? '');
      if (key === last || events >= 128) return;
      last = key;
      events++;
      try {
        emit({
          requestId,
          phase: event.phase,
          cancellable: safe,
          ...(event.recoveryId ? { recoveryId: event.recoveryId } : {}),
        });
      } catch {}
    },
    check() {
      if (granted) throw stopped();
    },
    enterCritical(phase) {
      gate.check();
      locked = true;
      gate.progress({ phase, cancellable: false });
    },
    leaveCritical() {
      gate.check();
      locked = false;
      safe = true;
    },
    authorize() {
      if (ended) return { granted: false, reason: 'completed' };
      if (locked || !safe) return { granted: false, reason: 'finishing' };
      granted = true;
      safe = false;
      gate.progress({ phase: 'stopping', cancellable: false });
      return { granted: true };
    },
    finish() {
      ended = true;
      safe = false;
    },
    get granted() {
      return granted;
    },
  };
  return gate;
}
