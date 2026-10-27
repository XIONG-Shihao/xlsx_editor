// Test oracle captured before performance migration; never imported by production.
import type { WorkbookEdit, WorkbookSheet } from '@office-editors/xlsx-editor/contracts'
export function colName(index: number): string {
  let name = ''
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26))
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name
  return name
}
export function addressPosition(ref: string): { row: number; col: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(ref)!
  return {
    col: [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
    row: Number(match[2]) - 1
  }
}
export function applyGridEdit(
  sheet: WorkbookSheet,
  edit: WorkbookEdit
): WorkbookSheet {
  const cells = { ...sheet.cells }
  if (edit.cell)
    cells[edit.cell] = {
      ...cells[edit.cell],
      ...(edit.value !== undefined
        ? { value: edit.value, formula: undefined }
        : {}),
      ...(edit.formula !== undefined ? { formula: edit.formula } : {}),
      ...(edit.bold !== undefined ? { bold: edit.bold } : {}),
      ...(edit.italic !== undefined ? { italic: edit.italic } : {}),
      ...(edit.numberFormat !== undefined
        ? {
            format: { general: 0, number: 2, percent: 10, date: 14, text: 49 }[
              edit.numberFormat
            ]
          }
        : {})
    }
  if (edit.operation) {
    const row = edit.operation.endsWith('Row'),
      remove = edit.operation.startsWith('delete'),
      index = edit.index ?? 0
    for (const key of Object.keys(cells)) delete cells[key]
    for (const [key, value] of Object.entries(sheet.cells)) {
      const p = addressPosition(key),
        n = row ? p.row : p.col
      if (remove && n === index) continue
      if (n >= index) {
        if (row) p.row += remove ? -1 : 1
        else p.col += remove ? -1 : 1
      }
      cells[`${colName(p.col)}${p.row + 1}`] = value
    }
  }
  return {
    ...sheet,
    cells,
    rows: Math.max(
      sheet.rows,
      ...Object.keys(cells).map((ref) => addressPosition(ref).row + 1)
    ),
    columns: Math.max(
      sheet.columns,
      ...Object.keys(cells).map((ref) => addressPosition(ref).col + 1)
    )
  }
}
function formulaDisplay(
  sheet: WorkbookSheet,
  ref: string,
  seen = new Set<string>()
): string | number {
  if (seen.has(ref)) return '#CYCLE!'
  const cell = sheet.cells[ref]
  if (!cell?.formula)
    return typeof cell?.value === 'number'
      ? cell.value
      : String(cell?.value ?? '')
  seen.add(ref)
  let formula = cell.formula.replace(/^=/, '').toUpperCase()
  try {
    formula = formula.replace(
      /(SUM|AVERAGE|MIN|MAX|COUNT)\(([A-Z]+\d+):([A-Z]+\d+)\)/g,
      (_all, fn: string, from: string, to: string) => {
        const a = addressPosition(from),
          b = addressPosition(to)
        if ((b.row - a.row + 1) * (b.col - a.col + 1) > 10000) throw Error()
        const values: number[] = []
        for (let r = a.row; r <= b.row; r++)
          for (let c = a.col; c <= b.col; c++) {
            const val = formulaDisplay(
              sheet,
              `${colName(c)}${r + 1}`,
              new Set(seen)
            )
            if (typeof val === 'number') values.push(val)
          }
        return String(
          fn === 'COUNT'
            ? values.length
            : fn === 'SUM'
              ? values.reduce((a, b) => a + b, 0)
              : fn === 'AVERAGE'
                ? values.reduce((a, b) => a + b, 0) / values.length
                : fn === 'MIN'
                  ? Math.min(...values)
                  : Math.max(...values)
        )
      }
    )
    formula = formula.replace(
      /\$?([A-Z]+)\$?(\d+)/g,
      (_m, col: string, row: string) => {
        const val = formulaDisplay(sheet, `${col}${row}`, new Set(seen))
        if (val === '') return '0'
        if (typeof val !== 'number') throw Error()
        return String(val)
      }
    )
    const tokens =
      formula.match(/\d+(?:\.\d+)?(?:E[+-]?\d+)?|[()+\-*/^]/g) ?? []
    if (tokens.join('') !== formula.replace(/\s/g, '')) throw Error()
    let i = 0
    function primary(): number {
      const token = tokens[i++]
      if (token === '-') return -primary()
      if (token === '+') return primary()
      if (token === '(') {
        const value = add()
        if (tokens[i++] !== ')') throw Error()
        return value
      }
      const value = Number(token)
      if (!Number.isFinite(value)) throw Error()
      return value
    }
    function power(): number {
      let value = primary()
      if (tokens[i] === '^') {
        i++
        value = Math.pow(value, power())
      }
      return value
    }
    function multiply(): number {
      let value = power()
      while (tokens[i] === '*' || tokens[i] === '/') {
        const op = tokens[i++],
          right = power()
        value = op === '*' ? value * right : value / right
      }
      return value
    }
    function add(): number {
      let value = multiply()
      while (tokens[i] === '+' || tokens[i] === '-') {
        const op = tokens[i++],
          right = multiply()
        value = op === '+' ? value + right : value - right
      }
      return value
    }
    const value = add()
    if (i !== tokens.length || !Number.isFinite(value)) throw Error()
    return value
  } catch {
    return 'Recalculate externally'
  }
}
export function displayValue(sheet: WorkbookSheet, ref: string): string {
  const cell = sheet.cells[ref],
    value = formulaDisplay(sheet, ref)
  if (typeof value !== 'number' || !Number.isFinite(value)) return String(value)
  const format = cell?.format
  if (
    [9, 10].includes(Number(format)) ||
    (typeof format === 'string' && format.includes('%'))
  ) {
    const decimals = format === 9 ? 0 : 2
    return `${(value * 100).toFixed(decimals)}%`
  }
  if (
    [14, 15, 16, 17, 22].includes(Number(format)) ||
    (typeof format === 'string' &&
      /[yd]/i.test(format.replace(/"[^"]*"|\\./g, '')))
  ) {
    if (!sheet.date1904 && Math.floor(value) === 60) return '1900-02-29'
    const days = sheet.date1904 ? value : value < 60 ? value + 1 : value
    const date = new Date(
      (sheet.date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)) +
        days * 86400000
    )
    return Number.isFinite(date.getTime())
      ? date.toISOString().slice(0, 10)
      : String(value)
  }
  if (format === 2 || format === '0.00') return value.toFixed(2)
  return String(value)
}

export { formulaDisplay }
