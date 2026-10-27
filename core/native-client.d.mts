import type { MessagePort } from 'node:worker_threads';
export interface WorkbookNativeClient {
  call(input: Record<string, unknown>): Promise<any>;
  close(error?: Error): void;
}
export function createWorkbookNativeClient(binary: string): WorkbookNativeClient;
/** Parse native frames inside a worker using a host-owned raw-byte relay. */
export function createWorkbookNativePortClient(port: MessagePort): WorkbookNativeClient;
