/** Grid input policy shared by the renderer and the trusted import owner. */
export function parseCellInput(text, literal = false) {
  if (literal) return { value: text };
  if (text.startsWith("'")) return { value: text.slice(1) };
  if (text === '') return { value: null };
  if (text.startsWith('=')) return { formula: text.slice(1) };
  if (/^(TRUE|FALSE)$/i.test(text)) return { value: text.toUpperCase() === 'TRUE' };
  if (/^[+-]?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(text)) {
    const number = Number(text);
    if (
      Number.isFinite(number) &&
      (!Number.isInteger(number) || Number.isSafeInteger(number)) &&
      text.replace(/[^0-9]/g, '').length <= 15
    )
      return { value: number };
  }
  return { value: text };
}
