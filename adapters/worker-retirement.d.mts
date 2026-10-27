import type {Worker} from 'node:worker_threads';
export function workerRetirement(worker:Worker):()=>Promise<void>;
