/** Resolve active source/output independently of the caller's working directory. */
import { mkdirSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { nativeWorkbookBinary } from '../../core/native-runtime.mjs';
const root = fileURLToPath(new URL('../../', import.meta.url));
export const componentPath = (relative) => path.join(root, relative);
export const currentCore = () => componentPath('core');
export const currentBinary = () => nativeWorkbookBinary();
export function benchmarkDirectory(phase) {
  const directory = componentPath(path.join('.build/runs/bench', phase));
  mkdirSync(directory, { recursive: true });
  return directory;
}
