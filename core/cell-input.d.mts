export function parseCellInput(
  text: string,
  literal?: boolean
): { value: string | number | boolean | null; formula?: undefined } | { formula: string; value?: undefined };
