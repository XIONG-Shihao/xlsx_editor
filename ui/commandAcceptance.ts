/** Strict registered commands keep the initiating runtime and accepted head.
 * Ordinary point edits use the host's existing FIFO instead. */
export interface OwnerCommandBinding {
  readonly projectId: string;
  readonly entryId: string;
  readonly documentId: string;
  readonly generation: number;
  readonly kind: string;
  readonly sessionId: string;
  readonly state: string;
}
export interface OwnerCommandCurrent extends OwnerCommandBinding {
  readonly editable: boolean;
  readonly active?: boolean;
}
export class OwnerCommandRetired extends Error {
  readonly code = 'OWNER_COMMAND_RETIRED';
  constructor() {
    super('The initiating workbook or accepted revision changed before this command.');
    this.name = 'OwnerCommandRetired';
  }
}
export class OwnerCommandCancelled extends Error {
  readonly code = 'OWNER_COMMAND_CANCELLED';
  constructor() {
    super('The workbook command was cancelled. Accepted recovery data was kept.');
    this.name = 'OwnerCommandCancelled';
  }
}
const runtimeKeys = ['projectId', 'entryId', 'documentId', 'generation', 'kind', 'sessionId'] as const;
export function captureOwnerCommandBinding(binding: OwnerCommandBinding): OwnerCommandBinding {
  const strings = ['projectId', 'entryId', 'documentId', 'kind', 'sessionId', 'state'] as const;
  if (
    !binding ||
    strings.some((key) => typeof binding[key] !== 'string' || !binding[key].length) ||
    !Number.isSafeInteger(binding.generation) ||
    binding.generation < 0
  )
    throw new OwnerCommandRetired();
  return Object.freeze({
    projectId: binding.projectId,
    entryId: binding.entryId,
    documentId: binding.documentId,
    generation: binding.generation,
    kind: binding.kind,
    sessionId: binding.sessionId,
    state: binding.state,
  });
}
/** A completed command may advance the head without changing its owner. */
export function ownerCommandRuntimeMatches(
  captured: OwnerCommandBinding,
  current: OwnerCommandCurrent | null
): boolean {
  return !!current && runtimeKeys.every((key) => captured[key] === current[key]);
}
export function ownerCommandBindingMatches(
  captured: OwnerCommandBinding,
  current: OwnerCommandCurrent | null
): boolean {
  return ownerCommandRuntimeMatches(captured, current) && captured.state === current!.state;
}
export function assertOwnerCommandBinding(captured: OwnerCommandBinding, current: OwnerCommandCurrent | null): void {
  if (!ownerCommandBindingMatches(captured, current) || current!.active === false) throw new OwnerCommandRetired();
  if (!current!.editable) throw Error('Workbook is busy or read-only.');
}
export async function queueCapturedOwnerCommand<T>(
  captured: OwnerCommandBinding,
  callbacks: {
    current: () => OwnerCommandCurrent | null;
    enqueue: (work: () => Promise<T>) => Promise<T>;
    apply: (binding: OwnerCommandBinding) => Promise<T>;
  }
): Promise<T> {
  assertOwnerCommandBinding(captured, callbacks.current());
  return callbacks.enqueue(async () => {
    assertOwnerCommandBinding(captured, callbacks.current());
    // Acceptance remains true after a tab or target moves. The host separately
    // adopts the result only into the original runtime, never into its successor.
    return callbacks.apply(captured);
  });
}
/** Cancellation can return a recovered owner instead of rejecting the IPC.
 * Preserve that recovered state, but never report it as command acceptance. */
export function acceptRecoveredCommand<T extends { operationCancelled?: boolean; cancelled?: boolean }>(
  result: T | null | undefined,
  update: (value: T) => void
): void {
  if (!result) throw new OwnerCommandRetired();
  update(result);
  if (result.operationCancelled === true || result.cancelled === true) throw new OwnerCommandCancelled();
}
