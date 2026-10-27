/** Reconstructible whole-rule summaries. Visible results live separately; an
 * oversized summary may render once without becoming a retained cache entry. */
export class RangeSummaryCache {
  private entries = new Map<string, { sheet: string; range: string; values: string[]; bytes: number }>();
  private bytes = 0;
  private readonly maxBytes: number;
  private readonly maxEntries: number;
  constructor(maxBytes = 2 * 1024 * 1024, maxEntries = 128) {
    this.maxBytes = maxBytes;
    this.maxEntries = maxEntries;
  }
  get(sheet: string, range: string): string[] | undefined {
    const key = JSON.stringify([sheet, range]),
      entry = this.entries.get(key);
    if (!entry) return undefined;
    this.entries.delete(key);
    this.entries.set(key, entry);
    return entry.values;
  }
  set(sheet: string, range: string, values: string[]): void {
    const key = JSON.stringify([sheet, range]),
      prior = this.entries.get(key);
    if (prior) {
      this.entries.delete(key);
      this.bytes -= prior.bytes;
    }
    // Account UTF-16 strings plus conservative per-entry/value allowances.
    // This is a payload admission estimate, not a JavaScript heap cap.
    const bytes =
      128 +
      (key.length + sheet.length + range.length) * 2 +
      values.reduce((sum, value) => sum + 32 + value.length * 2, 0);
    if (bytes > this.maxBytes) return;
    this.entries.set(key, { sheet, range, values, bytes });
    this.bytes += bytes;
    while (this.bytes > this.maxBytes || this.entries.size > this.maxEntries) {
      const oldest = this.entries.keys().next().value!;
      this.bytes -= this.entries.get(oldest)!.bytes;
      this.entries.delete(oldest);
    }
  }
  diagnostics() {
    return {
      duplicateSheets: new Set([...this.entries.values()].map((entry) => entry.sheet)).size,
      retainedSummaries: this.entries.size,
      estimatedSummaryCacheBytes: this.bytes,
    };
  }
  diagnosticEntries(): ReadonlyMap<string, unknown> {
    return this.entries;
  }
}
