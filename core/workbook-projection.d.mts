export function projectWorkbookCommands<T>(
  view: T,
  commands: unknown[],
  engine: unknown
): Promise<{ view: T; edits: unknown[] }>;

export function projectCellCommands(
  input: any,
  commands: any[]
): { edits: any[]; changes: { sheet: string; cell: string; value: any }[] } | null;
