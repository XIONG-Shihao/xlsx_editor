export function nativeWorkbookBinary(): string;
export function nativeWorkbookIdentity(): {
  schema: string;
  platform: string;
  arch: string;
  sha256: string;
  scope: string;
} | null;
export function nativeWorkbookOperationAsync(input: Record<string, unknown>): Promise<any>;
export function nativeWorkbookOperation(input: Record<string, unknown>): any;
export function closeNativeWorkbookWorker(): void;
