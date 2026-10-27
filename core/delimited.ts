import { refusal } from './refusal-codes.mjs';
export interface DelimitedDocument {
  rows: string[][];
  delimiter: string;
  newline: string;
  terminalNewline: boolean;
  original: string;
  raw: string[][];
  originalRows: string[][];
}
/** Count separators outside quoted fields; prefer comma when evidence ties. */
export function detectDelimiter(text: string): ',' | ';' | '\t' {
  const counts = {
    ',': [] as number[],
    ';': [] as number[],
    '\t': [] as number[],
  };
  let quoted = false,
    current = { ',': 0, ';': 0, '\t': 0 };
  for (let i = 0; i < text.length && counts[','].length < 20; i++) {
    const c = text[i];
    if (c === '"') {
      if (quoted && text[i + 1] === '"') i++;
      else quoted = !quoted;
    } else if (!quoted && (c === ',' || c === ';' || c === '\t')) current[c]++;
    else if (!quoted && (c === '\n' || c === '\r')) {
      if (c === '\r' && text[i + 1] === '\n') i++;
      for (const delimiter of [',', ';', '\t'] as const) counts[delimiter].push(current[delimiter]);
      current = { ',': 0, ';': 0, '\t': 0 };
    }
  }
  for (const delimiter of [',', ';', '\t'] as const) if (current[delimiter]) counts[delimiter].push(current[delimiter]);
  let best: ',' | ';' | '\t' = ',',
    score = 0;
  for (const delimiter of [',', ';', '\t'] as const) {
    const rows = counts[delimiter].filter((n) => n > 0);
    const candidate = rows.length && rows.every((n) => n === rows[0]) ? rows.length * rows[0] : 0;
    if (candidate > score) {
      best = delimiter;
      score = candidate;
    }
  }
  return best;
}
export function parseDelimited(text: string, delimiter: ',' | ';' | '\t'): DelimitedDocument {
  const rows: string[][] = [],
    raw: string[][] = [];
  let row: string[] = [],
    rawRow: string[] = [],
    value = '',
    field = '',
    quoted = false,
    closedQuote = false;
  const newline = text.match(/\r\n|\r|\n/)?.[0] ?? '\n';
  const finishCell = (): void => {
    row.push(value);
    rawRow.push(field);
    value = '';
    field = '';
    closedQuote = false;
  };
  const finishRow = (): void => {
    finishCell();
    rows.push(row);
    raw.push(rawRow);
    row = [];
    rawRow = [];
  };
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      field += c;
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += text[++i];
          value += '"';
        } else {
          quoted = false;
          closedQuote = true;
        }
      } else value += c;
    } else if (c === '"' && field === '') {
      quoted = true;
      field += c;
    } else if (c === delimiter) finishCell();
    else if (c === '\n' || c === '\r') {
      finishRow();
      if (c === '\r' && text[i + 1] === '\n') i++;
    } else {
      if (closedQuote) throw refusal('csv-malformed', 'Malformed delimited file: text follows a closing quote');
      value += c;
      field += c;
    }
  }
  if (quoted) throw refusal('csv-malformed', 'Malformed delimited file: unterminated quote');
  const terminalNewline = /[\r\n]$/.test(text);
  if (field || value || row.length || !rows.length) finishRow();
  return {
    rows,
    delimiter,
    newline,
    terminalNewline,
    original: text,
    raw,
    originalRows: rows.map((row) => [...row]),
  };
}
export function writeDelimited(document: DelimitedDocument, rows: string[][]): string {
  if (JSON.stringify(rows) === JSON.stringify(document.originalRows)) return document.original;
  return (
    rows
      .map((row, r) =>
        row
          .map((value, c) => {
            if (value === document.originalRows[r]?.[c]) return document.raw[r][c];
            return value.includes(document.delimiter) || /["\r\n]/.test(value) || document.raw[r]?.[c]?.startsWith('"')
              ? `"${value.replaceAll('"', '""')}"`
              : value;
          })
          .join(document.delimiter)
      )
      .join(document.newline) + (document.terminalNewline ? document.newline : '')
  );
}
