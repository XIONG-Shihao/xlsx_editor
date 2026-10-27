export function calculateIndexedWorkbook(
  input: any,
  native: (input: any) => Promise<any>
): Promise<{ version: number; inputCells: number; formulaCells: number; mode?: string; reason?: string }>;
