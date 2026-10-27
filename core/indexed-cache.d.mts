export interface IndexedCacheOptions {
  directory?: string;
  maxBytes?: number;
  maxEntries?: number;
  signal?: AbortSignal;
}
export interface IndexedCacheManifest {
  schema: number;
  revision: string;
  identity: string;
  version: number;
  bytes: number;
  reuseCalculation: boolean;
  checkpoint: boolean;
  files: Array<{ name: string; bytes: number; sha256: string }>;
}
export interface IndexedCachePublication {
  checkpoint?: unknown;
  workingBytes?: number;
  reuseCalculation?: boolean;
  prepareSnapshot?: (root: string) => Promise<unknown>;
}
export interface IndexedCache {
  revision: string;
  invalidate(): Promise<void>;
  restore(target: string, maxSessionBytes?: number): Promise<IndexedCacheManifest | null>;
  publish(
    sourceRoot: string,
    version: number,
    options?: IndexedCachePublication
  ): Promise<{ published: boolean; existing?: boolean; bytes?: number } | null>;
}
export function fileDigest(file: string, options?: { signal?: AbortSignal }): Promise<string>;
export function indexedImplementationIdentity(binary: string): Promise<string>;
export function indexedCache(source: string, identity: string, options?: IndexedCacheOptions): Promise<IndexedCache>;
export function indexedCacheUsage(
  options?: IndexedCacheOptions
): Promise<{ bytes: number; entries: number; staging: number; maxBytes: number; maxEntries: number }>;
