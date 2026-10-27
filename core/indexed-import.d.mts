import type { PasteDescriptor, PastePage } from './paste-protocol.mjs';
export function applyIndexedImport(options: {
  descriptor: PasteDescriptor;
  pages: AsyncIterable<PastePage>;
  current: any;
  native: (request: any) => Promise<any>;
  checkCancelled?: () => void;
  progress?: (value: { phase: 'applying-paste'; completed: number; total: number }) => void;
}): Promise<{
  working: any;
  version: number;
  view: any;
  changedRanges: { sheet: string; range: string }[];
  contentChanged: true;
  formulaChanged: boolean;
  formulaDelta: number;
  editCount: number;
  journal: PasteDescriptor;
  diagnostics: { peakPageCells: number; peakReadPositions: number; patchCalls: number };
}>;
