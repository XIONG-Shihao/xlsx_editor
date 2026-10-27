/** Optional, content-free diagnostics owned by a test/profiling host. Normal
 * startup installs no observer and does not walk workbook data. */
export type WorkbookDiagnosticKind = 'viewport' | 'renderer-worker';
export type WorkbookDiagnosticCounters = Record<string, number>;
export interface WorkbookDiagnosticsObserver {
  register(kind: WorkbookDiagnosticKind, read: () => WorkbookDiagnosticCounters): () => void;
}

export function registerWorkbookDiagnostics(
  kind: WorkbookDiagnosticKind,
  read: () => WorkbookDiagnosticCounters
): (() => void) | undefined {
  const observer = (globalThis as typeof globalThis & { __platformWorkbookDiagnostics?: WorkbookDiagnosticsObserver })
    .__platformWorkbookDiagnostics;
  if (!observer) return undefined;
  try {
    const release = observer.register(kind, () => {
      try {
        return Object.fromEntries(
          Object.entries(read())
            .filter(([key, value]) => /^[a-zA-Z][a-zA-Z0-9_]{0,63}$/.test(key) && Number.isFinite(value))
            .slice(0, 32)
        );
      } catch {
        return { diagnosticError: 1 };
      }
    });
    return () => {
      try {
        release();
      } catch {
        /* Diagnostics cannot affect disposal. */
      }
    };
  } catch {
    return () => {};
  }
}

/** Estimate reachable scalar payload, not JS heap/allocator usage. Count UTF-16
 * strings, numbers (8 bytes), booleans (1 byte) and typed backing data, excluding
 * object/Map overhead. Shared objects are counted once. Work and retained
 * traversal state are capped; truncated=1 means the estimate is a lower bound. */
export function diagnosticPayload(value: unknown, limit = 100_000): WorkbookDiagnosticCounters {
  const budget = Number.isSafeInteger(limit) && limit > 0 ? Math.min(limit, 100_000) : 100_000;
  const seen = new Set<object>(),
    stack: unknown[] = [value];
  let estimatedPayloadBytes = 0,
    visitedValues = 0,
    truncated = 0;
  const push = (item: unknown): boolean => {
    if (stack.length + visitedValues >= budget) {
      truncated = 1;
      return false;
    }
    stack.push(item);
    return true;
  };
  while (stack.length) {
    const item = stack.pop();
    visitedValues++;
    if (typeof item === 'string') {
      estimatedPayloadBytes += item.length * 2;
      continue;
    }
    if (typeof item === 'number') {
      estimatedPayloadBytes += 8;
      continue;
    }
    if (typeof item === 'boolean') {
      estimatedPayloadBytes++;
      continue;
    }
    if (!item || typeof item !== 'object' || seen.has(item)) continue;
    seen.add(item);
    if (ArrayBuffer.isView(item)) {
      estimatedPayloadBytes += item.byteLength;
      continue;
    }
    if (item instanceof ArrayBuffer) {
      estimatedPayloadBytes += item.byteLength;
      continue;
    }
    if (item instanceof Map) {
      for (const [key, entry] of item) {
        if (!push(key) || !push(entry)) break;
      }
    } else if (item instanceof Set) {
      for (const entry of item) if (!push(entry)) break;
    } else {
      for (const key in item) {
        if (!Object.prototype.hasOwnProperty.call(item, key)) continue;
        if (!push(key) || !push((item as Record<string, unknown>)[key])) break;
      }
    }
  }
  return { estimatedPayloadBytes, visitedValues, visitedObjects: seen.size, truncated };
}
